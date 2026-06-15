using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using B = DocumentFormat.OpenXml.Bibliography;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>The bibliography (b:) schema namespace; the customXml Sources part roots here, and CITATION fields reference it.</summary>
    private const string BibliographyNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

    /// <summary>The source kinds an agent passes in props.kind, mapped to the b:SourceType vocabulary Word stores.</summary>
    private static readonly Dictionary<string, string> SourceKinds = new(StringComparer.OrdinalIgnoreCase)
    {
        ["book"] = "Book",
        ["journalArticle"] = "JournalArticle",
        ["website"] = "InternetSite",
        ["report"] = "Report",
    };

    /// <summary>The citation styles a bibliography understands, mapped to Word's bundled style xsl filenames.</summary>
    private static readonly Dictionary<string, string> BibliographyStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["APA"] = "APASixthEditionOfficeOnline.xsl",
        ["MLA"] = "MLASeventhEditionOfficeOnline.xsl",
        ["Chicago"] = "ChicagoFifteenthEditionOfficeOnline.xsl",
    };

    // ================================================================== add source

    /// <summary>
    /// <c>{"op":"add","path":"/sources","type":"source","props":{"tag":"Smith2020","kind":"book","author":"Smith, John","title":"…","year":2020,…}}</c>:
    /// writes one bibliography source into the document's Sources store (the
    /// customXml part Word reads its bibliography from, rooted at <c>b:Sources</c>).
    /// The tag is the stable key citations reference; re-adding the same tag
    /// replaces it in place. Author is "Last, First" (or a corporate name).
    /// </summary>
    private static object ApplyAddSource(WordprocessingDocument doc, EditOp op)
    {
        if (!IsSourcesPath(op.Path))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type source targets /sources, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/sources\",\"type\":\"source\",\"props\":{\"tag\":\"Smith2020\",\"kind\":\"book\",\"title\":\"…\"}}.",
                candidates: ["/sources"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author"); // not the batch-author metadata here — re-read below as a content prop

        var tag = SourceProp(op.Props, "tag");
        if (string.IsNullOrWhiteSpace(tag))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type source needs props.tag (the citation key).",
                "Pass a short unique tag, e.g. {\"tag\":\"Smith2020\"} — citations reference the source by this tag.");
        }

        var kindArg = SourceProp(op.Props, "kind") ?? "book";
        if (!SourceKinds.TryGetValue(kindArg, out var sourceType))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown source kind '{kindArg}'.",
                $"Use kind {string.Join(", ", SourceKinds.Keys)}.",
                candidates: [.. SourceKinds.Keys]);
        }

        var sources = EnsureSourcesRoot(doc);
        var existing = FindSource(sources, tag);
        var refreshed = existing is not null;
        var source = BuildSource(tag, sourceType, op.Props);
        if (existing is not null)
        {
            existing.InsertAfterSelf(source);
            existing.Remove();
        }
        else
        {
            sources.AppendChild(source);
        }

        WriteSourcesPart(doc, sources);

        return new
        {
            op = "add",
            type = "source",
            path = SourcePath(tag),
            tag,
            kind = kindArg,
            replaced = refreshed ? true : (bool?)null,
        };
    }

    /// <summary>Builds one b:Source element from the add props (only the fields the kind uses are written).</summary>
    private static B.Source BuildSource(string tag, string sourceType, JsonObject? props)
    {
        var source = new B.Source(
            new B.SourceType(sourceType),
            // A stable GUID keeps Word from treating re-imports as distinct sources.
            new B.GuidString("{" + DeterministicGuid(tag).ToString().ToUpperInvariant() + "}"),
            new B.Tag(tag));

        if (SourceProp(props, "title") is { Length: > 0 } title)
        {
            source.AppendChild(new B.Title(title));
        }

        if (SourceProp(props, "year") is { Length: > 0 } year)
        {
            source.AppendChild(new B.Year(year));
        }

        if (SourceProp(props, "publisher") is { Length: > 0 } publisher)
        {
            source.AppendChild(new B.Publisher(publisher));
        }

        if (SourceProp(props, "journal") is { Length: > 0 } journal)
        {
            source.AppendChild(new B.JournalName(journal));
        }

        if (SourceProp(props, "url") is { Length: > 0 } url)
        {
            source.AppendChild(new B.UrlString(url));
        }

        if (SourceProp(props, "author") is { Length: > 0 } author)
        {
            source.AppendChild(BuildAuthorList(author));
        }

        return source;
    }

    /// <summary>An "Last, First" author becomes a b:Person; anything without a comma is a corporate author.</summary>
    private static B.AuthorList BuildAuthorList(string author)
    {
        var comma = author.IndexOf(',', StringComparison.Ordinal);
        if (comma < 0)
        {
            return new B.AuthorList(new B.Author(new B.Corporate(author.Trim())));
        }

        var last = author[..comma].Trim();
        var first = author[(comma + 1)..].Trim();
        var person = new B.Person(new B.Last(last));
        if (first.Length > 0)
        {
            person.AppendChild(new B.First(first));
        }

        return new B.AuthorList(new B.Author(new B.NameList(person)));
    }

    // ================================================================== get / remove source

    /// <summary>get /source[@tag=Smith2020]: the stored source props.</summary>
    private static (string Path, Dictionary<string, object?> Props) GetSourceProperties(WordprocessingDocument doc, DocPath path)
    {
        var tag = RequireSourceTag(path);
        var sources = ReadSourcesRoot(doc);
        var source = sources is null ? null : FindSource(sources, tag);
        if (source is null)
        {
            throw SourceNotFound(doc, tag);
        }

        return (SourcePath(tag), SourceShape(source));
    }

    /// <summary>remove /source[@tag=Smith2020]: drops the source from the store (existing citations keep their cached label).</summary>
    private static object ApplyRemoveSource(WordprocessingDocument doc, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        var tag = RequireSourceTag(path);
        var sources = ReadSourcesRoot(doc) ?? throw SourceNotFound(doc, tag);
        var source = FindSource(sources, tag) ?? throw SourceNotFound(doc, tag);

        source.Remove();
        WriteSourcesPart(doc, sources);
        return new { op = "remove", path = SourcePath(tag), type = "source", tag };
    }

    /// <summary>read --view sources: every stored source with its props, in document order.</summary>
    private static object SourcesView(WordprocessingDocument doc)
    {
        var sources = ReadSourcesRoot(doc);
        var list = sources is null
            ? []
            : sources.Elements<B.Source>()
                .Select(s => new { path = SourcePath(TagOf(s) ?? string.Empty), properties = SourceShape(s) })
                .ToList();

        return new { view = "sources", count = list.Count, sources = list };
    }

    /// <summary>The get/list shape of one source: tag, kind, and the populated bibliographic fields.</summary>
    private static Dictionary<string, object?> SourceShape(B.Source source)
    {
        var sourceType = source.GetFirstChild<B.SourceType>()?.Text;
        var kind = SourceKinds.FirstOrDefault(kv => kv.Value.Equals(sourceType, StringComparison.Ordinal)).Key;
        return new Dictionary<string, object?>
        {
            ["tag"] = TagOf(source),
            ["kind"] = kind ?? sourceType,
            ["title"] = source.GetFirstChild<B.Title>()?.Text,
            ["author"] = AuthorLabel(source),
            ["year"] = source.GetFirstChild<B.Year>()?.Text,
            ["publisher"] = source.GetFirstChild<B.Publisher>()?.Text,
            ["journal"] = source.GetFirstChild<B.JournalName>()?.Text,
            ["url"] = source.GetFirstChild<B.UrlString>()?.Text,
        };
    }

    // ================================================================== add citation

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"citation","props":{"source":"Smith2020","pages"?,"suppressAuthor"?}}</c>:
    /// appends a CITATION field referencing the source, with a cached rendered
    /// label (e.g. "(Smith, 2020)") so the document reads before Word recomputes.
    /// Unknown source tag → invalid_args naming the existing tags as candidates.
    /// </summary>
    private static object ApplyAddCitation(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var sourceTag = SourceProp(op.Props, "source");
        if (string.IsNullOrWhiteSpace(sourceTag))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type citation needs props.source (the tag of the source to cite).",
                "Pass {\"source\":\"Smith2020\"}; add the source first with {\"op\":\"add\",\"path\":\"/sources\",\"type\":\"source\",…}.");
        }

        var sources = ReadSourcesRoot(doc);
        var source = sources is null ? null : FindSource(sources, sourceTag);
        if (source is null)
        {
            throw SourceNotFound(doc, sourceTag);
        }

        var target = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (target.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Citations are appended to paragraphs, not '{target.Type}'.",
                target.Type is "tc" or "header" or "footer"
                    ? $"Target the paragraph inside it: {target.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[2].");
        }

        var pages = SourceProp(op.Props, "pages");
        var suppressAuthor = props["suppressAuthor"] is { } sa && IsTrue(sa);
        var label = CitationLabel(source, pages, suppressAuthor);
        var field = BuildCitationField(sourceTag, pages, suppressAuthor, label);
        DeclareDocumentW14Ignorable(doc); // CITATION fields carry a w:fldSimple; keep mc/w14 ready alongside other fields
        paragraph.Append(field);

        return new
        {
            op = "add",
            type = "citation",
            path = target.CanonicalPath,
            source = sourceTag,
            cached = label,
            note = "Word refreshes the citation when fields update; the cached label is what shows until then.",
        };
    }

    /// <summary>A CITATION field (w:fldSimple) referencing the source tag, with the rendered label as its cached result.</summary>
    private static OpenXmlElement[] BuildCitationField(string sourceTag, string? pages, bool suppressAuthor, string label)
    {
        var instruction = new StringBuilder(" CITATION ").Append(sourceTag).Append(' ');
        if (pages is { Length: > 0 })
        {
            instruction.Append("\\p ").Append(Quote(pages)).Append(' ');
        }

        if (suppressAuthor)
        {
            instruction.Append("\\n ");
        }

        var field = new SimpleField(new Run(NewText(label))) { Instruction = instruction.ToString() };
        return [field];
    }

    /// <summary>A cached "(Author, Year)" label (page/suffix appended), so the citation reads without Word recompute.</summary>
    private static string CitationLabel(B.Source source, string? pages, bool suppressAuthor)
    {
        var year = source.GetFirstChild<B.Year>()?.Text;
        var author = suppressAuthor ? null : AuthorSurname(source);

        var inner = new StringBuilder();
        if (author is { Length: > 0 })
        {
            inner.Append(author);
            if (year is { Length: > 0 })
            {
                inner.Append(", ");
            }
        }

        if (year is { Length: > 0 })
        {
            inner.Append(year);
        }

        if (pages is { Length: > 0 })
        {
            if (inner.Length > 0)
            {
                inner.Append(", p. ");
            }

            inner.Append(pages);
        }

        return inner.Length > 0 ? $"({inner})" : "(citation)";
    }

    // ================================================================== bibliography

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"bibliography","props":{"style":"APA"?}}</c>:
    /// a BIBLIOGRAPHY field plus one pre-rendered static paragraph per cited
    /// source (styled), so it reads without Word recompute. The exact numbering
    /// and format finalize when Word opens or refreshes the field — surfaced as a
    /// <c>bibliography_cached</c> warning, like the TOC. Re-running replaces it.
    /// </summary>
    private static object ApplyAddBibliography(WordprocessingDocument doc, string file, EditOp op, EditSession session)
    {
        if (!IsBodyPath(op.Path))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type bibliography targets /body, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/body\",\"type\":\"bibliography\",\"props\":{\"style\":\"APA\"}}.",
                candidates: ["/body"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        var styleArg = SourceProp(op.Props, "style") ?? "APA";
        if (!BibliographyStyles.ContainsKey(styleArg))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown bibliography style '{styleArg}'.",
                $"Use style {string.Join(", ", BibliographyStyles.Keys)}.",
                candidates: [.. BibliographyStyles.Keys]);
        }

        var sources = ReadSourcesRoot(doc);
        // Only sources that are actually cited belong in the bibliography; when no
        // citations exist yet, fall back to every stored source so the block is not empty.
        var citedTags = CitedSourceTags(doc);
        var entries = (sources?.Elements<B.Source>() ?? [])
            .Where(s => citedTags.Count == 0 || (TagOf(s) is { } t && citedTags.Contains(t)))
            .ToList();

        // Apply the selected style to the Sources store so Word renders consistently.
        if (sources is not null)
        {
            sources.SelectedStyle = "\\" + BibliographyStyles[styleArg];
            sources.StyleName = styleArg;
            WriteSourcesPart(doc, sources);
        }

        var body = GetBody(doc, file);
        var existing = EnumerateBibliographies(doc).FirstOrDefault();
        var replaced = existing is not null;
        existing?.Remove();

        EnsureBibliographyStyles(doc);
        var block = BuildBibliographyBlock(entries, styleArg);
        if (body.Elements<SectionProperties>().FirstOrDefault() is { } sectPr)
        {
            sectPr.InsertBeforeSelf(block);
        }
        else
        {
            body.AppendChild(block);
        }

        if (!session.Warnings.Any(w => w.Code == WarningCodes.BibliographyCached))
        {
            session.Warnings.Add(new Warning(
                WarningCodes.BibliographyCached,
                "Bibliography entries were pre-rendered from the stored sources so the document reads immediately. " +
                "Word finalizes the exact ordering and format when it opens the file (or on field refresh, F9)."));
        }

        return new
        {
            op = "add",
            type = "bibliography",
            path = BibliographyPath,
            style = styleArg,
            entries = entries.Count,
            replaced = replaced ? true : (bool?)null,
        };
    }

    /// <summary>remove /bibliography[1]: drops the sdt block (the sources store is untouched).</summary>
    private static object ApplyRemoveBibliography(WordprocessingDocument doc, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        if (path.Segments is not [{ Name: "bibliography" }])
        {
            throw BibliographyNotFound(doc);
        }

        var block = EnumerateBibliographies(doc).FirstOrDefault() ?? throw BibliographyNotFound(doc);
        block.Remove();
        return new { op = "remove", path = BibliographyPath, type = "bibliography" };
    }

    /// <summary>
    /// The bibliography sdt block: a BIBLIOGRAPHY complex field, then one styled
    /// paragraph per source pre-rendered as a static entry (Author. (Year). Title.
    /// Publisher/Journal. URL). Word keeps these between the field's separate and
    /// end markers; we put the field markers on the first/last entry paragraphs.
    /// </summary>
    private static SdtBlock BuildBibliographyBlock(List<B.Source> sources, string style)
    {
        var content = new SdtContentBlock();
        content.AppendChild(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(NewText("References"))));

        var fieldOpeners = new OpenXmlElement[]
        {
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" BIBLIOGRAPHY ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        };

        if (sources.Count == 0)
        {
            var lone = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Bibliography" }));
            lone.Append(fieldOpeners);
            lone.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            content.AppendChild(lone);
        }

        for (var i = 0; i < sources.Count; i++)
        {
            var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Bibliography" }));
            if (i == 0)
            {
                paragraph.Append(fieldOpeners);
            }

            paragraph.AppendChild(new Run(NewText(BibliographyEntry(sources[i], style))));

            if (i == sources.Count - 1)
            {
                paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            }

            content.AppendChild(paragraph);
        }

        return new SdtBlock(
            new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Bibliographies" },
                    new DocPartUnique())),
            content);
    }

    /// <summary>One pre-rendered static bibliography line for a source (a readable approximation of the chosen style).</summary>
    private static string BibliographyEntry(B.Source source, string style)
    {
        var author = AuthorLabel(source);
        var year = source.GetFirstChild<B.Year>()?.Text;
        var title = source.GetFirstChild<B.Title>()?.Text;
        var publisher = source.GetFirstChild<B.Publisher>()?.Text;
        var journal = source.GetFirstChild<B.JournalName>()?.Text;
        var url = source.GetFirstChild<B.UrlString>()?.Text;

        var parts = new List<string>();
        if (author is { Length: > 0 })
        {
            parts.Add(author + ".");
        }

        if (year is { Length: > 0 })
        {
            parts.Add($"({year}).");
        }

        if (title is { Length: > 0 })
        {
            parts.Add(title + ".");
        }

        if (journal is { Length: > 0 })
        {
            parts.Add(journal + ".");
        }

        if (publisher is { Length: > 0 })
        {
            parts.Add(publisher + ".");
        }

        if (url is { Length: > 0 })
        {
            parts.Add(url);
        }

        return parts.Count > 0 ? string.Join(' ', parts) : "(bibliography entry)";
    }

    /// <summary>The Bibliography paragraph style (hanging indent), defined on demand; Heading1 for the "References" title.</summary>
    private static void EnsureBibliographyStyles(WordprocessingDocument doc)
    {
        EnsureStyleDefined(doc, "Heading1");
        var styles = EnsureStylesRoot(doc);
        if (FindStyle(styles, "Bibliography") is null)
        {
            styles.AppendChild(new Style(
                new StyleName { Val = "Bibliography" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new Indentation { Left = "720", Hanging = "720" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Bibliography",
            });
        }
    }

    // ================================================================== store

    /// <summary>The customXml part rooted at b:Sources, or null when the document has no bibliography store yet.</summary>
    private static CustomXmlPart? FindSourcesPart(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.CustomXmlParts.FirstOrDefault(IsSourcesPart);

    private static bool IsSourcesPart(CustomXmlPart part)
    {
        try
        {
            // Own the part stream explicitly: XmlReader.Create(Stream) does not close
            // its input, and leaving the part stream open in update mode blocks the
            // next GetStream on the same part.
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = System.Xml.XmlReader.Create(stream);
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    return reader.LocalName == "Sources" && reader.NamespaceURI == BibliographyNamespace;
                }
            }
        }
        catch (Exception ex) when (ex is System.Xml.XmlException or InvalidOperationException or IOException)
        {
            return false;
        }

        return false;
    }

    /// <summary>Reads the live b:Sources element from the store part, or null when there is none.</summary>
    private static B.Sources? ReadSourcesRoot(WordprocessingDocument doc)
    {
        if (FindSourcesPart(doc) is not { } part)
        {
            return null;
        }

        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        var xml = reader.ReadToEnd();
        if (string.IsNullOrWhiteSpace(xml))
        {
            return null;
        }

        return new B.Sources(xml);
    }

    /// <summary>Reads the existing b:Sources, or creates an empty one (not yet persisted) ready to receive a source.</summary>
    private static B.Sources EnsureSourcesRoot(WordprocessingDocument doc) =>
        ReadSourcesRoot(doc) ?? new B.Sources
        {
            SelectedStyle = "\\" + BibliographyStyles["APA"],
            StyleName = "APA",
        };

    /// <summary>
    /// Persists the b:Sources element into the customXml store part, creating the
    /// part on first write. Uses FeedData (not a fresh Create stream): in update
    /// mode a part stream cannot be reopened for Create after it has been read,
    /// and a single edit reads the store before it writes it back.
    /// </summary>
    private static void WriteSourcesPart(WordprocessingDocument doc, B.Sources sources)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part to attach a bibliography store to.",
            "Re-export the file from Word.");

        var part = FindSourcesPart(doc) ?? main.AddCustomXmlPart("application/xml");
        var bytes = new UTF8Encoding(false).GetBytes(sources.OuterXml);
        using var ms = new MemoryStream(bytes);
        part.FeedData(ms);
    }

    // ================================================================== helpers

    private static B.Source? FindSource(B.Sources sources, string tag) =>
        sources.Elements<B.Source>().FirstOrDefault(s =>
            string.Equals(TagOf(s), tag, StringComparison.Ordinal));

    private static string? TagOf(B.Source source) => source.GetFirstChild<B.Tag>()?.Text;

    /// <summary>The tags of every source referenced by a CITATION field in the body, in document order.</summary>
    private static HashSet<string> CitedSourceTags(WordprocessingDocument doc)
    {
        var tags = new HashSet<string>(StringComparer.Ordinal);
        if (doc.MainDocumentPart?.Document?.Body is not { } body)
        {
            return tags;
        }

        foreach (var field in body.Descendants<SimpleField>())
        {
            if (ParseCitationTag(field.Instruction?.Value) is { } tag)
            {
                tags.Add(tag);
            }
        }

        foreach (var code in body.Descendants<FieldCode>())
        {
            if (ParseCitationTag(code.Text) is { } tag)
            {
                tags.Add(tag);
            }
        }

        return tags;
    }

    /// <summary>The source tag inside a " CITATION Smith2020 \p … " instruction, or null when it is not a citation.</summary>
    private static string? ParseCitationTag(string? instruction)
    {
        if (instruction is null)
        {
            return null;
        }

        var tokens = instruction.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return tokens is [var head, var tag, ..] && head.Equals("CITATION", StringComparison.OrdinalIgnoreCase)
            ? tag
            : null;
    }

    /// <summary>The full author label "Last, First" (first author only) or a corporate name, for read-back and entries.</summary>
    private static string? AuthorLabel(B.Source source)
    {
        if (source.Descendants<B.Corporate>().FirstOrDefault()?.Text is { Length: > 0 } corporate)
        {
            return corporate;
        }

        var person = source.Descendants<B.Person>().FirstOrDefault();
        var last = person?.GetFirstChild<B.Last>()?.Text;
        var first = person?.GetFirstChild<B.First>()?.Text;
        if (last is null or { Length: 0 })
        {
            return null;
        }

        return first is { Length: > 0 } ? $"{last}, {first}" : last;
    }

    /// <summary>Just the surname (or corporate name) for the in-text "(Author, Year)" label.</summary>
    private static string? AuthorSurname(B.Source source)
    {
        if (source.Descendants<B.Corporate>().FirstOrDefault()?.Text is { Length: > 0 } corporate)
        {
            return corporate;
        }

        return source.Descendants<B.Person>().FirstOrDefault()?.GetFirstChild<B.Last>()?.Text;
    }

    private static AiofficeException SourceNotFound(WordprocessingDocument doc, string tag)
    {
        var sources = ReadSourcesRoot(doc);
        var tags = sources?.Elements<B.Source>().Select(TagOf).OfType<string>().ToList() ?? [];
        return new AiofficeException(
            ErrorCodes.InvalidArgs,
            tags.Count == 0
                ? $"No source is tagged '{tag}'; the document has no bibliography sources yet."
                : $"No source is tagged '{tag}'.",
            tags.Count == 0
                ? "Add the source first: {\"op\":\"add\",\"path\":\"/sources\",\"type\":\"source\",\"props\":{\"tag\":\"" + tag + "\",\"kind\":\"book\",\"title\":\"…\"}}."
                : "Pick an existing tag, or add this source first with add --type source.",
            candidates: tags.Count > 0 ? tags : [tag]);
    }

    private static AiofficeException BibliographyNotFound(WordprocessingDocument doc)
    {
        _ = doc;
        return new AiofficeException(
            ErrorCodes.InvalidPath,
            "This document has no bibliography.",
            "Add one with {\"op\":\"add\",\"path\":\"/body\",\"type\":\"bibliography\",\"props\":{\"style\":\"APA\"}}.",
            candidates: [BibliographyPath]);
    }

    /// <summary>The top-level bibliography sdt gallery blocks, in document order.</summary>
    private static List<SdtBlock> EnumerateBibliographies(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body is { } body
            ? [.. body.Elements<SdtBlock>().Where(IsBibliographyBlock)]
            : [];

    private static bool IsBibliographyBlock(SdtBlock sdt) =>
        sdt.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
            == "Bibliographies";

    private const string BibliographyPath = "/bibliography[1]";

    private static string SourcePath(string tag) => $"/source[@tag={tag}]";

    private static bool IsSourcesPath(string path) =>
        path is "/sources" || DocPath.Parse(path).Segments is [{ Name: "sources", Index: null, Id: null }];

    private static bool IsBodyPath(string path)
    {
        var segments = DocPath.Parse(path).Segments;
        return segments is [{ Name: "body", Index: null, Id: null }];
    }

    /// <summary>The tag from a /source[@tag=…] path; throws invalid_args when the path is not a source address.</summary>
    private static string RequireSourceTag(DocPath path)
    {
        if (path.Segments is [{ Name: "source", Id: { } tag }])
        {
            return tag;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{path.ToCanonicalString()}' is not a source path.",
            "Use /source[@tag=Smith2020] (the tag you gave when adding the source).");
    }

    /// <summary>A source/citation prop as a trimmed string (null when absent or blank).</summary>
    private static string? SourceProp(JsonObject? props, string name) =>
        props?[name] is { } node && NodeToString(node).Trim() is { Length: > 0 } value ? value : null;

    /// <summary>A stable GUID derived from the tag, so re-adding the same source keeps Word's identity consistent.</summary>
    private static Guid DeterministicGuid(string tag)
    {
        var bytes = System.Security.Cryptography.MD5.HashData(Encoding.UTF8.GetBytes("aioffice:source:" + tag));
        return new Guid(bytes);
    }

    private static string Quote(string value) => "\"" + value.Replace("\"", string.Empty, StringComparison.Ordinal) + "\"";
}
