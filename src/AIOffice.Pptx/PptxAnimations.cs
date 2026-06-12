using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>One animation effect on a slide (the par hosting an effect cTn), 1-based index in document order.</summary>
internal sealed record AnimationView(
    int Index,
    P.ParallelTimeNode Par,
    P.CommonTimeNode TimeNode,
    string Effect,
    string Trigger,
    uint? ShapeId,
    string? Duration,
    string? Delay,
    string? Direction);

/// <summary>
/// Entrance animations (p:timing): add via
/// <c>{"op":"add","path":"/slide[i]/shape[@id=N]","type":"animation","props":{...}}</c>,
/// addressed as /slide[i]/animation[k] for get/remove. The timing tree follows
/// PowerPoint's shape — tmRoot par → mainSeq → one par per click group, effects
/// nested two pars deep with the standard preset class/id/subtype and a
/// style.visibility set behavior (plus animEffect/anim per effect).
/// </summary>
internal static class PptxAnimations
{
    /// <summary>The effects aioffice can add. Everything else is unsupported_feature.</summary>
    public static readonly IReadOnlyList<string> Effects = ["appear", "fade", "flyIn", "wipe"];

    /// <summary>The triggers aioffice understands.</summary>
    public static readonly IReadOnlyList<string> Triggers = ["click", "afterPrevious", "withPrevious"];

    private static readonly IReadOnlyList<string> Directions = ["left", "right", "top", "bottom"];

    private static readonly IReadOnlyList<string> AddProps =
        ["effect", "trigger", "duration", "delay", "direction"];

    private const int AppearPresetId = 1;
    private const int FlyInPresetId = 2;
    private const int FadePresetId = 10;
    private const int WipePresetId = 22;

    // ----- add -------------------------------------------------------------------

    /// <summary>Adds one entrance animation targeting a shape; returns /slide[i]/animation[k].</summary>
    public static string Add(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        if (!address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add animation targets a shape, not '{address.Raw}'.",
                "Use the shape path as the target, e.g. {\"op\":\"add\",\"path\":\"/slide[1]/shape[@id=5]\"," +
                "\"type\":\"animation\",\"props\":{\"effect\":\"fade\"}}.");
        }

        if (address.ParagraphIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Paragraph-level animation is not supported yet.",
                "Animate the whole shape instead: target /slide[i]/shape[@id=N].");
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = PptxDoc.ResolveShape(slidePart, address);
        var spec = ParseProps(props);

        var slide = slidePart.Slide ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The slide part has no slide XML.",
            "The slide part is malformed; re-export the file or restore a snapshot.");

        var mainSequenceChildren = EnsureMainSequence(slide);
        var nextId = NextTimeNodeId(slide.Timing!);

        var groupChildren = spec.Trigger == "click" || LastGroupChildren(mainSequenceChildren) is null
            ? NewClickGroup(mainSequenceChildren, ref nextId)
            : LastGroupChildren(mainSequenceChildren)!;
        groupChildren.Append(BuildEffectPar(spec, view.Id, ref nextId));

        var index = List(slidePart).Count;
        return Units.Inv($"/slide[{address.SlideIndex}]/animation[{index}]");
    }

    private sealed record EffectSpec(string Effect, string Trigger, long DurationMs, long DelayMs, string Direction);

    private static EffectSpec ParseProps(JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown animation prop '{key}'.",
                    "Animation props: effect, trigger, duration, delay, direction.",
                    candidates: AddProps);
            }
        }

        var effect = Token(props, "effect") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "add animation needs the 'effect' prop.",
            "Pass one of appear, fade, flyIn or wipe, e.g. {\"effect\":\"fade\",\"trigger\":\"click\"}.",
            candidates: Effects);
        if (!Effects.Contains(effect, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Animation effect '{effect}' is not supported.",
                "Supported effects: appear, fade, flyIn, wipe. Pick the closest one and refine it in PowerPoint.",
                candidates: Effects);
        }

        effect = Effects.First(e => string.Equals(e, effect, StringComparison.OrdinalIgnoreCase));

        var trigger = Token(props, "trigger") ?? "click";
        if (!Triggers.Contains(trigger, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown animation trigger '{trigger}'.",
                "Use click (starts on click), afterPrevious or withPrevious.",
                candidates: Triggers);
        }

        trigger = Triggers.First(t => string.Equals(t, trigger, StringComparison.OrdinalIgnoreCase));

        var direction = Token(props, "direction");
        if (direction is not null)
        {
            if (effect is not ("flyIn" or "wipe"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop 'direction' does not apply to '{effect}'.",
                    "Only flyIn and wipe take a direction (left, right, top, bottom).");
            }

            if (!Directions.Contains(direction, StringComparer.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown animation direction '{direction}'.",
                    "Use left, right, top or bottom.",
                    candidates: Directions);
            }

            direction = Directions.First(d => string.Equals(d, direction, StringComparison.OrdinalIgnoreCase));
        }

        var durationMs = props.TryGetPropertyValue("duration", out var durationNode)
            ? ParseSecondsMs("duration", durationNode, max: 600)
            : 500;
        if (effect == "appear" && props.ContainsKey("duration"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Prop 'duration' does not apply to 'appear' (it is instantaneous).",
                "Drop duration, or use fade for a timed entrance.");
        }

        var delayMs = props.TryGetPropertyValue("delay", out var delayNode)
            ? ParseSecondsMs("delay", delayNode, max: 600, allowZero: true)
            : 0;

        return new EffectSpec(effect, trigger, durationMs, delayMs, direction ?? "bottom");
    }

    private static string? Token(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is not null
            ? J.ScalarText(node).Trim()
            : null;

    /// <summary>Duration/delay in ms. Accepts "0.5s", "500ms" or a plain number of seconds.</summary>
    private static long ParseSecondsMs(string key, JsonNode? node, double max, bool allowZero = false)
    {
        double? seconds = null;
        if (node is JsonValue value)
        {
            if (Units.TryNumber(value, out var plain))
            {
                seconds = plain;
            }
            else if (value.TryGetValue<string>(out var raw))
            {
                var text = raw.Trim().ToLowerInvariant();
                var (suffix, factor) = text switch
                {
                    _ when text.EndsWith("ms", StringComparison.Ordinal) => ("ms", 0.001),
                    _ when text.EndsWith("s", StringComparison.Ordinal) => ("s", 1.0),
                    _ => ("", 1.0),
                };
                var numberText = suffix.Length == 0 ? text : text[..^suffix.Length].Trim();
                if (double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    seconds = number * factor;
                }
            }
        }

        var floor = allowZero ? 0 : double.Epsilon;
        if (seconds is null || seconds < floor || seconds > max)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid {key}: {node?.ToJsonString() ?? "null"}",
                Units.Inv($"Use seconds like \"0.5s\" (or \"500ms\"); {key} must be {(allowZero ? "0 or more" : "positive")} and at most {max}s."));
        }

        return (long)Math.Round(seconds.Value * 1000);
    }

    // ----- timing-tree plumbing ----------------------------------------------------

    /// <summary>The mainSeq childTnLst, building the root timing skeleton when missing.</summary>
    private static P.ChildTimeNodeList EnsureMainSequence(P.Slide slide)
    {
        var mainSequence = slide.Timing?.Descendants<P.CommonTimeNode>()
            .FirstOrDefault(c => c.NodeType?.Value == P.TimeNodeValues.MainSequence);
        if (mainSequence is not null)
        {
            return mainSequence.GetFirstChild<P.ChildTimeNodeList>()
                ?? mainSequence.AppendChild(new P.ChildTimeNodeList());
        }

        var timing = slide.Timing;
        var nextId = timing is null ? 1u : NextTimeNodeId(timing);

        var mainSequenceChildren = new P.ChildTimeNodeList();
        var sequence = new P.SequenceTimeNode(
            new P.CommonTimeNode(mainSequenceChildren)
            {
                Id = nextId + 1,
                Duration = "indefinite",
                NodeType = P.TimeNodeValues.MainSequence,
            },
            new P.PreviousConditionList(new P.Condition(new P.TargetElement(new P.SlideTarget()))
            {
                Event = P.TriggerEventValues.OnPrevious,
                Delay = "0",
            }),
            new P.NextConditionList(new P.Condition(new P.TargetElement(new P.SlideTarget()))
            {
                Event = P.TriggerEventValues.OnNext,
                Delay = "0",
            }))
        {
            Concurrent = true,
            NextAction = P.NextActionValues.Seek,
        };

        if (timing?.TimeNodeList?.Elements<P.ParallelTimeNode>().FirstOrDefault()?.CommonTimeNode is { } root &&
            root.NodeType?.Value == P.TimeNodeValues.TmingRoot)
        {
            // Foreign timing without a main sequence: hook ours into the existing root.
            (root.GetFirstChild<P.ChildTimeNodeList>() ?? root.AppendChild(new P.ChildTimeNodeList()))
                .Append(sequence);
            return mainSequenceChildren;
        }

        slide.Timing = new P.Timing(new P.TimeNodeList(new P.ParallelTimeNode(
            new P.CommonTimeNode(new P.ChildTimeNodeList(sequence))
            {
                Id = nextId,
                Duration = "indefinite",
                Restart = P.TimeNodeRestartValues.Never,
                NodeType = P.TimeNodeValues.TmingRoot,
            })));
        return mainSequenceChildren;
    }

    private static uint NextTimeNodeId(P.Timing timing)
    {
        var max = 0u;
        foreach (var node in timing.Descendants<P.CommonTimeNode>())
        {
            if (node.Id?.Value is { } id && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    /// <summary>The inner childTnLst of the last click group, or null when no group exists yet.</summary>
    private static P.ChildTimeNodeList? LastGroupChildren(P.ChildTimeNodeList mainSequenceChildren) =>
        mainSequenceChildren.Elements<P.ParallelTimeNode>().LastOrDefault()?
            .CommonTimeNode?.GetFirstChild<P.ChildTimeNodeList>()?
            .Elements<P.ParallelTimeNode>().FirstOrDefault()?
            .CommonTimeNode?.GetFirstChild<P.ChildTimeNodeList>();

    /// <summary>Appends a new click group (outer par delay=indefinite, inner par delay=0) and returns its inner childTnLst.</summary>
    private static P.ChildTimeNodeList NewClickGroup(P.ChildTimeNodeList mainSequenceChildren, ref uint nextId)
    {
        var inner = new P.ChildTimeNodeList();
        mainSequenceChildren.Append(new P.ParallelTimeNode(new P.CommonTimeNode(
            new P.StartConditionList(new P.Condition { Delay = "indefinite" }),
            new P.ChildTimeNodeList(new P.ParallelTimeNode(new P.CommonTimeNode(
                new P.StartConditionList(new P.Condition { Delay = "0" }),
                inner)
            {
                Id = nextId + 1,
                Fill = P.TimeNodeFillValues.Hold,
            })))
        {
            Id = nextId,
            Fill = P.TimeNodeFillValues.Hold,
        }));
        nextId += 2;
        return inner;
    }

    /// <summary>The effect par: preset cTn + visibility set + per-effect behaviors.</summary>
    private static P.ParallelTimeNode BuildEffectPar(EffectSpec spec, uint shapeId, ref uint nextId)
    {
        var (presetId, presetSubtype) = spec.Effect switch
        {
            "appear" => (AppearPresetId, 0),
            "fade" => (FadePresetId, 0),
            "flyIn" => (FlyInPresetId, spec.Direction switch
            {
                "left" => 8,
                "right" => 2,
                "top" => 1,
                _ => 4, // bottom
            }),
            _ => (WipePresetId, spec.Direction switch // "wipe"
            {
                "left" => 2,
                "right" => 8,
                "top" => 4,
                _ => 1, // bottom
            }),
        };

        var children = new P.ChildTimeNodeList(BuildVisibilitySet(shapeId, ref nextId));
        switch (spec.Effect)
        {
            case "fade":
                children.Append(BuildAnimateEffect(shapeId, "fade", spec.DurationMs, ref nextId));
                break;
            case "wipe":
                var filter = spec.Direction switch
                {
                    "left" => "wipe(right)",
                    "right" => "wipe(left)",
                    "top" => "wipe(down)",
                    _ => "wipe(up)",
                };
                children.Append(BuildAnimateEffect(shapeId, filter, spec.DurationMs, ref nextId));
                break;
            case "flyIn":
                var (fromX, fromY) = spec.Direction switch
                {
                    "left" => ("0-#ppt_w/2", "#ppt_y"),
                    "right" => ("1+#ppt_w/2", "#ppt_y"),
                    "top" => ("#ppt_x", "0-#ppt_h/2"),
                    _ => ("#ppt_x", "1+#ppt_h/2"), // bottom
                };
                children.Append(BuildAnimate(shapeId, "ppt_x", fromX, "#ppt_x", spec.DurationMs, ref nextId));
                children.Append(BuildAnimate(shapeId, "ppt_y", fromY, "#ppt_y", spec.DurationMs, ref nextId));
                break;
            default:
                break; // appear: visibility only
        }

        var effectNode = new P.CommonTimeNode(
            new P.StartConditionList(new P.Condition
            {
                Delay = spec.DelayMs.ToString(CultureInfo.InvariantCulture),
            }),
            children)
        {
            Id = nextId,
            PresetId = presetId,
            PresetClass = P.TimeNodePresetClassValues.Entrance,
            PresetSubtype = presetSubtype,
            Fill = P.TimeNodeFillValues.Hold,
            GroupId = 0,
            NodeType = spec.Trigger switch
            {
                "withPrevious" => P.TimeNodeValues.WithEffect,
                "afterPrevious" => P.TimeNodeValues.AfterEffect,
                _ => P.TimeNodeValues.ClickEffect,
            },
        };
        nextId++;
        return new P.ParallelTimeNode(effectNode);
    }

    private static P.SetBehavior BuildVisibilitySet(uint shapeId, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode(new P.StartConditionList(new P.Condition { Delay = "0" }))
            {
                Id = nextId++,
                Duration = "1",
                Fill = P.TimeNodeFillValues.Hold,
            },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) }),
            new P.AttributeNameList(new P.AttributeName("style.visibility"))),
        new P.ToVariantValue(new P.StringVariantValue { Val = "visible" }));

    private static P.AnimateEffect BuildAnimateEffect(uint shapeId, string filter, long durationMs, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode { Id = nextId++, Duration = Inv(durationMs) },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) })))
    {
        Transition = P.AnimateEffectTransitionValues.In,
        Filter = filter,
    };

    private static P.Animate BuildAnimate(uint shapeId, string attribute, string from, string to, long durationMs, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode { Id = nextId++, Duration = Inv(durationMs), Fill = P.TimeNodeFillValues.Hold },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) }),
            new P.AttributeNameList(new P.AttributeName(attribute))),
        new P.TimeAnimateValueList(
            new P.TimeAnimateValue(new P.VariantValue(new P.StringVariantValue { Val = from })) { Time = "0" },
            new P.TimeAnimateValue(new P.VariantValue(new P.StringVariantValue { Val = to })) { Time = "100000" }))
    {
        CalculationMode = P.AnimateBehaviorCalculateModeValues.Linear,
        ValueType = P.AnimateBehaviorValues.Number,
    };

    private static string Inv(long value) => value.ToString(CultureInfo.InvariantCulture);

    private static string Inv(uint value) => value.ToString(CultureInfo.InvariantCulture);

    // ----- read side -----------------------------------------------------------------

    /// <summary>All animation effects on a slide in document order, 1-based.</summary>
    public static List<AnimationView> List(SlidePart slidePart)
    {
        var views = new List<AnimationView>();
        var timing = slidePart.Slide?.Timing;
        if (timing is null)
        {
            return views;
        }

        foreach (var node in timing.Descendants<P.CommonTimeNode>())
        {
            var trigger = node.NodeType?.Value switch
            {
                { } t when t == P.TimeNodeValues.ClickEffect => "click",
                { } t when t == P.TimeNodeValues.WithEffect => "withPrevious",
                { } t when t == P.TimeNodeValues.AfterEffect => "afterPrevious",
                _ => null,
            };
            if (trigger is null || node.Parent is not P.ParallelTimeNode par)
            {
                continue;
            }

            var presetId = node.PresetId?.Value;
            var effect = presetId switch
            {
                AppearPresetId => "appear",
                FadePresetId => "fade",
                FlyInPresetId => "flyIn",
                WipePresetId => "wipe",
                { } other => Units.Inv($"preset{other}"), // foreign decks: truthful, not pretty
                null => "unknown",
            };

            var direction = effect switch
            {
                "flyIn" => node.PresetSubtype?.Value switch
                {
                    8 => "left", 2 => "right", 1 => "top", 4 => "bottom", _ => null,
                },
                "wipe" => node.PresetSubtype?.Value switch
                {
                    2 => "left", 8 => "right", 4 => "top", 1 => "bottom", _ => null,
                },
                _ => null,
            };

            views.Add(new AnimationView(
                views.Count + 1,
                par,
                node,
                effect,
                trigger,
                FirstShapeTargetId(par),
                ReadDuration(par),
                ReadDelay(node),
                direction));
        }

        return views;
    }

    private static uint? FirstShapeTargetId(P.ParallelTimeNode par)
    {
        var spid = par.Descendants<P.ShapeTarget>().FirstOrDefault()?.ShapeId?.Value;
        return uint.TryParse(spid, NumberStyles.None, CultureInfo.InvariantCulture, out var id) ? id : null;
    }

    /// <summary>The first timed behavior duration under the effect ("0.5s"); null for appear-style effects.</summary>
    private static string? ReadDuration(P.ParallelTimeNode par)
    {
        foreach (var behavior in par.Descendants<P.CommonBehavior>())
        {
            if (behavior.Parent is P.SetBehavior)
            {
                continue; // the visibility set's 1ms pulse is not the user-facing duration
            }

            if (long.TryParse(
                behavior.GetFirstChild<P.CommonTimeNode>()?.Duration?.Value,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out var ms))
            {
                return FormatSeconds(ms);
            }
        }

        return null;
    }

    private static string? ReadDelay(P.CommonTimeNode node)
    {
        var delay = node.GetFirstChild<P.StartConditionList>()?.GetFirstChild<P.Condition>()?.Delay?.Value;
        return long.TryParse(delay, NumberStyles.None, CultureInfo.InvariantCulture, out var ms)
            ? FormatSeconds(ms)
            : null;
    }

    private static string FormatSeconds(long ms) =>
        (ms / 1000.0).ToString("0.###", CultureInfo.InvariantCulture) + "s";

    /// <summary>Resolves /slide[i]/animation[k] or throws invalid_path with candidates.</summary>
    public static AnimationView Resolve(SlidePart slidePart, PptxAddress address)
    {
        var animations = List(slidePart);
        var index = address.AnimationIndex!.Value;
        if (index >= 1 && index <= animations.Count)
        {
            return animations[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No animation {index} on slide {address.SlideIndex}; it has {animations.Count} animation(s)."),
            animations.Count > 0
                ? "Animation indices are 1-based per slide; run 'aioffice read <file> --view structure' to list them."
                : "Add one first: {\"op\":\"add\",\"path\":\"" + address.CanonicalSlidePath +
                  "/shape[@id=N]\",\"type\":\"animation\",\"props\":{\"effect\":\"fade\"}}.",
            candidates: [.. animations.Take(10).Select(a => Units.Inv($"{address.CanonicalSlidePath}/animation[{a.Index}]"))]);
    }

    /// <summary>The `get` projection for /slide[i]/animation[k].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = Resolve(slidePart, address);
        return Project(view, address.SlideIndex, slidePart);
    }

    /// <summary>The shared animation row (get and read --view structure).</summary>
    public static object Project(AnimationView view, int slideIndex, SlidePart slidePart)
    {
        string? targetPath = null;
        if (view.ShapeId is { } shapeId)
        {
            var exists = PptxDoc.Shapes(slidePart).Any(s => s.Id == shapeId);
            targetPath = exists ? Units.Inv($"/slide[{slideIndex}]/shape[@id={shapeId}]") : null;
        }

        return new
        {
            Path = Units.Inv($"/slide[{slideIndex}]/animation[{view.Index}]"),
            Slide = slideIndex,
            Index = view.Index,
            Target = targetPath,
            TargetShapeId = view.ShapeId,
            Effect = view.Effect,
            Trigger = view.Trigger,
            Duration = view.Duration,
            Delay = view.Delay,
            Direction = view.Direction,
        };
    }

    // ----- remove ---------------------------------------------------------------------

    /// <summary>remove /slide[i]/animation[k]: drops the effect par and prunes empty timing ancestors.</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = Resolve(slidePart, address);
        var slide = slidePart.Slide!;

        OpenXmlElement current = view.Par;
        while (true)
        {
            var container = current.Parent;
            current.Remove();

            if (container is P.TimeNodeList && !container.HasChildren)
            {
                slide.Timing = null; // an empty p:tnLst is schema-invalid; drop the whole tree
                break;
            }

            if (container is not P.ChildTimeNodeList childList || childList.HasChildren)
            {
                break;
            }

            // An empty childTnLst is schema-invalid: remove the par/seq that owns it and keep pruning.
            var owner = childList.Parent?.Parent; // childTnLst -> cTn -> par|seq
            if (owner is null)
            {
                childList.Remove();
                break;
            }

            current = owner;
        }

        return address.CanonicalAnimationPath;
    }
}
