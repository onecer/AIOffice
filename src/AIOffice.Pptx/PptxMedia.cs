using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace AIOffice.Pptx;

/// <summary>
/// Embedded slide media (video/audio). The wire form is a <c>p:pic</c> whose
/// <c>nvPr</c> carries an <c>a:videoFile</c>/<c>a:audioFile</c> (an r:link to a
/// <see cref="MediaDataPart"/>) plus a <c>p14:media</c> extension (an r:embed to
/// the same part), a <c>ppaction://media</c> click action, and a matching media
/// node in the slide's <c>p:timing</c> tree — exactly what PowerPoint writes for
/// "Insert → Video/Audio → from file". The host picture's shape id is the media's
/// stable id, so <c>/slide[i]/media[@id=N]</c> survives insertions and removals.
/// </summary>
internal static class PptxMedia
{
    /// <summary>The p14 media extension uri PowerPoint marks an embedded media picture with.</summary>
    private const string MediaExtensionUri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

    /// <summary>The click action that turns the picture into a media trigger.</summary>
    private const string MediaAction = "ppaction://media";

    private const string Office2010MainNs = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    private static readonly IReadOnlyList<string> AddPropKeys = ["src", "poster", "x", "y", "w", "h", "name", "autoplay"];

    /// <summary>One media object resolved on a slide: its host picture, the media data part and reference relationships.</summary>
    private sealed record MediaRef(int Ordinal, P.Picture Picture, MediaDataPart Part, uint HostId, bool IsVideo);

    /// <summary>Sniffed media identity: content type, file extension and whether it is video (vs audio).</summary>
    private readonly record struct MediaInfo(string ContentType, string Extension, bool IsVideo);

    // ---- add ----------------------------------------------------------------

    /// <summary>
    /// Embeds the sandbox-resolved <c>props.src</c> media file (video or audio) on
    /// the slide and returns the canonical media path (by the host picture's id).
    /// </summary>
    public static string Add(PresentationDocument document, SlidePart slidePart, int slideIndex, JsonObject? props, Workspace workspace)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown media prop '{key}'.",
                    "Media props: src (required), poster, x, y, w, h, name, autoplay. " +
                    "src points at an mp4/mov/m4a/mp3/wav inside the workspace.",
                    candidates: AddPropKeys);
            }
        }

        if (!props.TryGetPropertyValue("src", out var srcNode) || srcNode is null || J.ScalarText(srcNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add media requires props.src.",
                "Point src at a media file inside the workspace, e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"media\",\"props\":{\"src\":\"clip.mp4\"}}.");
        }

        // Sandbox first: a src outside the workspace is sandbox_denied, never read.
        var src = workspace.Resolve(J.ScalarText(srcNode).Trim(), mustExist: true);
        var info = Sniff(src);
        var bytes = File.ReadAllBytes(src);

        // props.poster (optional): a PNG/JPEG sandbox-resolved before any mutation.
        string? posterRel = null;
        if (props.TryGetPropertyValue("poster", out var posterNode) && posterNode is not null && J.ScalarText(posterNode).Trim().Length > 0)
        {
            var poster = workspace.Resolve(J.ScalarText(posterNode).Trim(), mustExist: true);
            posterRel = PptxImages.EmbedImagePart(slidePart, poster);
        }

        var autoplay = props.TryGetPropertyValue("autoplay", out var autoNode) && autoNode is not null && AsBool(autoNode);

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);
        var w = props.TryGetPropertyValue("w", out var wNode)
            ? Units.ParseLengthEmu("w", wNode)
            : Units.CmToEmu(info.IsVideo ? 16 : 4);
        var h = props.TryGetPropertyValue("h", out var hNode)
            ? Units.ParseLengthEmu("h", hNode)
            : Units.CmToEmu(info.IsVideo ? 9 : 4);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null && J.ScalarText(nameNode).Trim().Length > 0
            ? J.ScalarText(nameNode).Trim()
            : Path.GetFileName(src);

        // The verbatim payload lives in a MediaDataPart; FeedData copies the bytes
        // through with no re-encoding, so the media round-trips exactly.
        var mediaPart = document.CreateMediaDataPart(info.ContentType, info.Extension);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            mediaPart.FeedData(stream);
        }

        var fileRel = info.IsVideo
            ? slidePart.AddVideoReferenceRelationship(mediaPart).Id
            : slidePart.AddAudioReferenceRelationship(mediaPart).Id;
        var embedRel = slidePart.AddMediaReferenceRelationship(mediaPart).Id;

        tree.Append(BuildPicture(slidePart, id, name, info.IsVideo, fileRel, embedRel, posterRel, x, y, w, h));
        AppendTimingNode(slidePart, id, info.IsVideo, autoplay);

        return Units.Inv($"/slide[{slideIndex}]/media[@id={id}]");
    }

    /// <summary>Builds the p:pic that hosts the media (videoFile/audioFile + p14:media extension + click action).</summary>
    private static P.Picture BuildPicture(
        SlidePart slidePart, uint id, string name, bool isVideo, string fileRel, string embedRel, string? posterRel, long x, long y, long w, long h)
    {
        var nonVisualProps = new P.NonVisualDrawingProperties(
            new A.HyperlinkOnClick { Id = string.Empty, Action = MediaAction })
        {
            Id = id,
            Name = name,
        };

        var appProps = new P.ApplicationNonVisualDrawingProperties();
        appProps.Append(isVideo
            ? new A.VideoFromFile { Link = fileRel }
            : (OpenXmlElement)new A.AudioFromFile { Link = fileRel });

        var extension = new P.ApplicationNonVisualDrawingPropertiesExtension(
            new P14.Media { Embed = embedRel })
        {
            Uri = MediaExtensionUri,
        };
        extension.GetFirstChild<P14.Media>()!.AddNamespaceDeclaration("p14", Office2010MainNs);
        appProps.Append(new P.ApplicationNonVisualDrawingPropertiesExtensionList(extension));

        var blip = new A.Blip();
        if (posterRel is not null)
        {
            blip.Embed = posterRel;
        }

        return new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisualProps,
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                appProps),
            new P.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = w, Cy = h }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    /// <summary>
    /// Adds (or extends) the slide's p:timing tree with a media node targeting the
    /// host picture, so PowerPoint shows playback controls. autoplay starts the
    /// node on the slide's load instead of waiting for a click.
    /// </summary>
    private static void AppendTimingNode(SlidePart slidePart, uint shapeId, bool isVideo, bool autoplay)
    {
        var slide = slidePart.Slide ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The slide part has no slide XML.",
            "The slide part is malformed; re-export the file or restore a snapshot.");

        var mainSequence = EnsureTimingMainSequence(slide);
        var childList = mainSequence.GetFirstChild<P.ChildTimeNodeList>()
            ?? (P.ChildTimeNodeList)mainSequence.AppendChild(new P.ChildTimeNodeList());

        var startDelay = autoplay ? "0" : "indefinite";
        var commonMediaNode = new P.CommonMediaNode(
            new P.CommonTimeNode(
                new P.StartConditionList(new P.Condition { Delay = startDelay }))
            {
                Id = NextTimeNodeId(slide),
                Fill = P.TimeNodeFillValues.Hold,
                Display = false,
            },
            new P.TargetElement(new P.ShapeTarget { ShapeId = shapeId.ToString(System.Globalization.CultureInfo.InvariantCulture) }))
        {
            Volume = 50000,
        };

        childList.Append(isVideo
            ? new P.Video(commonMediaNode)
            : (OpenXmlElement)new P.Audio(commonMediaNode));
    }

    /// <summary>Ensures the slide has a p:timing root + a main sequence cTn and returns that sequence's cTn.</summary>
    private static P.CommonTimeNode EnsureTimingMainSequence(P.Slide slide)
    {
        var timing = slide.Timing;
        if (timing is null)
        {
            var seqNode = new P.SequenceTimeNode(
                new P.CommonTimeNode(new P.ChildTimeNodeList())
                {
                    Id = 2u,
                    Duration = "indefinite",
                    NodeType = P.TimeNodeValues.MainSequence,
                });
            timing = new P.Timing(
                new P.TimeNodeList(
                    new P.ParallelTimeNode(
                        new P.CommonTimeNode(new P.ChildTimeNodeList(seqNode))
                        {
                            Id = 1u,
                            Duration = "indefinite",
                            Restart = P.TimeNodeRestartValues.Never,
                            NodeType = P.TimeNodeValues.TmingRoot,
                        })));
            slide.Append(timing);
            return seqNode.GetFirstChild<P.CommonTimeNode>()!;
        }

        // A timing tree already exists (e.g. animations were added first): reuse
        // its main sequence, or graft one under the root par when none is present.
        var sequence = timing.Descendants<P.SequenceTimeNode>()
            .FirstOrDefault(s => s.GetFirstChild<P.CommonTimeNode>()?.NodeType?.Value == P.TimeNodeValues.MainSequence);
        if (sequence?.GetFirstChild<P.CommonTimeNode>() is { } existing)
        {
            return existing;
        }

        var rootChildList = timing.Descendants<P.ParallelTimeNode>().FirstOrDefault()?
            .GetFirstChild<P.CommonTimeNode>()?.GetFirstChild<P.ChildTimeNodeList>();
        var newSequence = new P.SequenceTimeNode(
            new P.CommonTimeNode(new P.ChildTimeNodeList())
            {
                Id = NextTimeNodeId(slide),
                Duration = "indefinite",
                NodeType = P.TimeNodeValues.MainSequence,
            });
        if (rootChildList is not null)
        {
            rootChildList.Append(newSequence);
        }
        else
        {
            // No usable root par: rebuild a minimal timing tree around the new sequence.
            timing.RemoveAllChildren();
            timing.Append(new P.TimeNodeList(
                new P.ParallelTimeNode(
                    new P.CommonTimeNode(new P.ChildTimeNodeList(newSequence))
                    {
                        Id = 1u,
                        Duration = "indefinite",
                        Restart = P.TimeNodeRestartValues.Never,
                        NodeType = P.TimeNodeValues.TmingRoot,
                    })));
        }

        return newSequence.GetFirstChild<P.CommonTimeNode>()!;
    }

    /// <summary>The next free p:cTn id on the slide (timing node ids must be unique within the tree).</summary>
    private static uint NextTimeNodeId(P.Slide slide)
    {
        uint max = 0;
        foreach (var node in slide.Timing?.Descendants<P.CommonTimeNode>() ?? [])
        {
            if (node.Id?.Value is { } id && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    // ---- list / get ---------------------------------------------------------

    /// <summary>The media objects on one slide, in document (ordinal) order.</summary>
    private static List<MediaRef> MediaOn(SlidePart slidePart)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var result = new List<MediaRef>();
        var ordinal = 0;
        foreach (var picture in tree.Elements<P.Picture>())
        {
            var nvPr = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            var video = nvPr?.GetFirstChild<A.VideoFromFile>();
            var audio = nvPr?.GetFirstChild<A.AudioFromFile>();
            var isVideo = video is not null;
            var fileRel = video?.Link?.Value ?? audio?.Link?.Value;
            if (fileRel is null)
            {
                continue;
            }

            // The reference relationship resolves to the embedded MediaDataPart.
            var reference = slidePart.DataPartReferenceRelationships
                .FirstOrDefault(r => r.Id == fileRel);
            if (reference?.DataPart is not MediaDataPart mediaPart)
            {
                continue;
            }

            ordinal++;
            var hostId = PptxDoc.NonVisualProps(picture)?.Id?.Value ?? 0;
            result.Add(new MediaRef(ordinal, picture, mediaPart, hostId, isVideo));
        }

        return result;
    }

    /// <summary>The media objects on one slide, projected for the structure view.</summary>
    public static List<object> SlideMedia(SlidePart slidePart, int slideIndex) =>
        MediaOn(slidePart).Select(m => Project(slidePart, slideIndex, m)).ToList();

    /// <summary>"video"/"audio" when the picture hosts embedded media; null otherwise (the render hook).</summary>
    public static string? MediaKindOf(P.Picture picture)
    {
        var nvPr = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        if (nvPr?.GetFirstChild<A.VideoFromFile>() is not null)
        {
            return "video";
        }

        return nvPr?.GetFirstChild<A.AudioFromFile>() is not null ? "audio" : null;
    }

    /// <summary>The number of embedded media objects on one slide (used by convert to report the loss).</summary>
    public static int SlideMediaCount(SlidePart slidePart) => MediaOn(slidePart).Count;

    /// <summary>The detail projection of one addressed media object (kind, src name, geometry).</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var media = Resolve(slidePart, address);
        return Project(slidePart, address.SlideIndex, media);
    }

    private static object Project(SlidePart slidePart, int slideIndex, MediaRef media)
    {
        var geometry = PptxDoc.Geometry(media.Picture);
        var name = PptxDoc.NonVisualProps(media.Picture)?.Name?.Value;
        return new
        {
            Path = Units.Inv($"/slide[{slideIndex}]/media[@id={media.HostId}]"),
            ShapePath = Units.Inv($"/slide[{slideIndex}]/shape[@id={media.HostId}]"),
            Slide = slideIndex,
            Id = media.HostId,
            Kind = media.IsVideo ? "video" : "audio",
            Src = string.IsNullOrWhiteSpace(name) ? Units.Inv($"Media {media.HostId}") : name,
            MediaType = media.Part.ContentType,
            Autoplay = AutoplayOf(slidePart, media.HostId),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
        };
    }

    /// <summary>True when the media node targeting the shape starts on load (delay 0) rather than on click.</summary>
    private static bool AutoplayOf(SlidePart slidePart, uint shapeId)
    {
        var shapeIdText = shapeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        foreach (var node in slidePart.Slide?.Timing?.Descendants<P.CommonMediaNode>() ?? [])
        {
            if (node.GetFirstChild<P.TargetElement>()?.ShapeTarget?.ShapeId?.Value == shapeIdText &&
                node.CommonTimeNode?.StartConditionList?.GetFirstChild<P.Condition>()?.Delay?.Value == "0")
            {
                return true;
            }
        }

        return false;
    }

    // ---- remove -------------------------------------------------------------

    /// <summary>Removes the addressed media (its host picture, the timing node and the media data part when orphaned).</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var media = Resolve(slidePart, address);
        var canonical = Units.Inv($"/slide[{address.SlideIndex}]/media[@id={media.HostId}]");

        RemoveTimingNode(slidePart, media.HostId);
        DeleteMediaPartsFor(slidePart, media.Picture);
        media.Picture.Remove();
        return canonical;
    }

    /// <summary>
    /// Drops the reference relationships (and the underlying media data part when
    /// nothing else references it) for a media-hosting picture being removed, so
    /// removing the picture does not orphan the payload.
    /// </summary>
    public static void DeleteMediaPartsFor(SlidePart slidePart, P.Picture picture)
    {
        var nvPr = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        var fileRel = nvPr?.GetFirstChild<A.VideoFromFile>()?.Link?.Value
            ?? nvPr?.GetFirstChild<A.AudioFromFile>()?.Link?.Value;
        var embedRel = nvPr?.GetFirstChild<P.ApplicationNonVisualDrawingPropertiesExtensionList>()?
            .Descendants<P14.Media>().FirstOrDefault()?.Embed?.Value;

        MediaDataPart? dataPart = null;
        foreach (var relId in new[] { fileRel, embedRel })
        {
            if (relId is null)
            {
                continue;
            }

            var reference = slidePart.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == relId);
            if (reference is not null)
            {
                dataPart ??= reference.DataPart as MediaDataPart;
                slidePart.DeleteReferenceRelationship(reference);
            }
        }

        // Delete the shared media data part only when no other slide references it.
        if (dataPart is not null && !PartStillReferenced(slidePart, dataPart))
        {
            slidePart.OpenXmlPackage.DeletePart(dataPart);
        }
    }

    private static bool PartStillReferenced(SlidePart removingFrom, MediaDataPart dataPart)
    {
        foreach (var part in removingFrom.OpenXmlPackage.GetPartsOfType<SlidePart>())
        {
            foreach (var reference in part.DataPartReferenceRelationships)
            {
                if (ReferenceEquals(reference.DataPart, dataPart))
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Removes the media node targeting the shape from the slide's timing tree and
    /// drops the whole p:timing when nothing animated/media remains — an empty
    /// childTnLst is invalid, so the tree is pruned rather than left hollow.
    /// </summary>
    private static void RemoveTimingNode(SlidePart slidePart, uint shapeId)
    {
        var timing = slidePart.Slide?.Timing;
        if (timing is null)
        {
            return;
        }

        var shapeIdText = shapeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        foreach (var node in timing.Descendants<OpenXmlElement>().Where(e => e is P.Video or P.Audio).ToList())
        {
            var common = (node as P.Video)?.CommonMediaNode ?? (node as P.Audio)?.CommonMediaNode;
            if (common?.GetFirstChild<P.TargetElement>()?.ShapeTarget?.ShapeId?.Value == shapeIdText)
            {
                node.Remove();
            }
        }

        // If no behavior nodes remain anywhere in the tree, the timing tree is
        // hollow (an empty childTnLst fails validation): drop it entirely.
        var hasBehavior = timing.Descendants<OpenXmlElement>().Any(e =>
            e is P.Video or P.Audio ||
            e.LocalName is "set" or "anim" or "animClr" or "animEffect" or "animMotion" or "animRot" or "animScale" or "cmd");
        if (!hasBehavior)
        {
            timing.Remove();
        }
    }

    // ---- resolution ---------------------------------------------------------

    private static MediaRef Resolve(SlidePart slidePart, PptxAddress address)
    {
        var media = MediaOn(slidePart);
        MediaRef? match = address.MediaId is { } id
            ? media.FirstOrDefault(m => m.HostId == id)
            : media.FirstOrDefault(m => m.Ordinal == address.MediaOrdinal);

        if (match is not null)
        {
            return match;
        }

        var what = address.MediaId is { } mid ? $"id {mid}" : $"index {address.MediaOrdinal}";
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No media object with {what} on slide {address.SlideIndex}; it has {media.Count} media object(s).",
            "Run 'aioffice read <file> --view structure' to list slide media and their canonical paths.",
            candidates: [.. media.Take(10).Select(m => Units.Inv($"/slide[{address.SlideIndex}]/media[@id={m.HostId}]"))]);
    }

    // ---- sniff --------------------------------------------------------------

    /// <summary>Identifies the media kind from file headers (mp4/mov video, m4a/mp3/wav audio); extension is a fallback.</summary>
    private static MediaInfo Sniff(string src)
    {
        var bytes = ReadHeader(src, 16);

        // ISO-BMFF (mp4/mov/m4a): bytes 4..7 spell "ftyp"; the major brand picks video vs audio.
        if (bytes.Length >= 12 && bytes[4] == 'f' && bytes[5] == 't' && bytes[6] == 'y' && bytes[7] == 'p')
        {
            var brand = System.Text.Encoding.ASCII.GetString(bytes, 8, 4);
            if (brand.StartsWith("M4A", StringComparison.Ordinal))
            {
                return new MediaInfo("audio/mp4", "m4a", IsVideo: false);
            }

            if (brand.StartsWith("qt", StringComparison.Ordinal))
            {
                return new MediaInfo("video/quicktime", "mov", IsVideo: true);
            }

            // isom/mp42/avc1/iso2/etc. — treat as mp4 video.
            return new MediaInfo("video/mp4", "mp4", IsVideo: true);
        }

        // RIFF WAVE audio: "RIFF"...."WAVE".
        if (bytes.Length >= 12 && bytes[0] == 'R' && bytes[1] == 'I' && bytes[2] == 'F' && bytes[3] == 'F' &&
            bytes[8] == 'W' && bytes[9] == 'A' && bytes[10] == 'V' && bytes[11] == 'E')
        {
            return new MediaInfo("audio/wav", "wav", IsVideo: false);
        }

        // MP3: an ID3 tag ("ID3") or an MPEG audio frame sync (0xFF 0xEx/0xFx).
        if (bytes.Length >= 3 && bytes[0] == 'I' && bytes[1] == 'D' && bytes[2] == '3')
        {
            return new MediaInfo("audio/mpeg", "mp3", IsVideo: false);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFF && (bytes[1] & 0xE0) == 0xE0)
        {
            return new MediaInfo("audio/mpeg", "mp3", IsVideo: false);
        }

        return ByExtension(src) ?? throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"'{Path.GetFileName(src)}' is not a recognized media file (sniffed by header).",
            "Embed an mp4 or mov video, or an m4a, mp3 or wav audio file; convert other formats first " +
            "(e.g. 'ffmpeg -i input.ext output.mp4').");
    }

    private static MediaInfo? ByExtension(string src) =>
        Path.GetExtension(src).ToLowerInvariant() switch
        {
            ".mp4" => new MediaInfo("video/mp4", "mp4", IsVideo: true),
            ".mov" => new MediaInfo("video/quicktime", "mov", IsVideo: true),
            ".m4a" => new MediaInfo("audio/mp4", "m4a", IsVideo: false),
            ".mp3" => new MediaInfo("audio/mpeg", "mp3", IsVideo: false),
            ".wav" => new MediaInfo("audio/wav", "wav", IsVideo: false),
            _ => null,
        };

    private static byte[] ReadHeader(string file, int count)
    {
        using var stream = File.OpenRead(file);
        var buffer = new byte[Math.Min(count, (int)Math.Min(stream.Length, int.MaxValue))];
        var read = 0;
        while (read < buffer.Length)
        {
            var n = stream.Read(buffer, read, buffer.Length - read);
            if (n == 0)
            {
                break;
            }

            read += n;
        }

        return read == buffer.Length ? buffer : buffer[..read];
    }

    private static bool AsBool(JsonNode node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property 'autoplay' is not a boolean: {node.ToJsonString()}",
            "Use true or false.");
    }
}
