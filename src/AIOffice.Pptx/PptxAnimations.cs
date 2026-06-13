using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>One animation effect on a slide (the par hosting an effect cTn), 1-based index in document order.</summary>
internal sealed record AnimationView(
    int Index,
    P.ParallelTimeNode Par,
    P.CommonTimeNode TimeNode,
    string Class,
    string Effect,
    string Trigger,
    uint? ShapeId,
    string? Duration,
    string? Delay,
    string? Direction);

/// <summary>
/// Shape animations (p:timing): add via
/// <c>{"op":"add","path":"/slide[i]/shape[@id=N]","type":"animation","props":{...}}</c>,
/// addressed as /slide[i]/animation[k] for get/remove. The timing tree follows
/// PowerPoint's shape — tmRoot par → mainSeq → one par per click group, effects
/// nested two pars deep with the standard preset class/id/subtype. Entrances
/// carry a style.visibility "visible" set, exits hide the shape at the end, and
/// emphasis effects animate in place (scale/rotation/color behaviors).
/// </summary>
internal static class PptxAnimations
{
    /// <summary>The entrance effects aioffice can add.</summary>
    public static readonly IReadOnlyList<string> EntranceEffects = ["appear", "fade", "flyIn", "wipe"];

    /// <summary>The emphasis effects aioffice can add.</summary>
    public static readonly IReadOnlyList<string> EmphasisEffects = ["pulse", "grow", "spin", "colorPulse"];

    /// <summary>The exit effects aioffice can add.</summary>
    public static readonly IReadOnlyList<string> ExitEffects = ["fadeOut", "flyOut", "wipeOut"];

    /// <summary>Every effect aioffice can add. Everything else is unsupported_feature.</summary>
    public static readonly IReadOnlyList<string> Effects =
        [.. EntranceEffects, .. EmphasisEffects, .. ExitEffects];

    /// <summary>The triggers aioffice understands.</summary>
    public static readonly IReadOnlyList<string> Triggers = ["click", "afterPrevious", "withPrevious"];

    private static readonly IReadOnlyList<string> Directions = ["left", "right", "top", "bottom"];

    private static readonly IReadOnlyList<string> AddProps =
        ["effect", "trigger", "duration", "delay", "direction", "color"];

    private const int AppearPresetId = 1;
    private const int FlyInPresetId = 2;
    private const int FadePresetId = 10;
    private const int WipePresetId = 22;

    private const int GrowPresetId = 6; // Grow/Shrink
    private const int SpinPresetId = 8;
    private const int PulsePresetId = 35;
    private const int ColorPulsePresetId = 36;

    /// <summary>360 degrees in OOXML 60000ths-of-a-degree units.</summary>
    private const int FullTurn = 21_600_000;

    /// <summary>Pulse scales to 106%, grow to 150% (percent * 1000, PowerPoint's defaults).</summary>
    private const int PulseScale = 106_000;
    private const int GrowScale = 150_000;

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

    // ----- set (retime) ----------------------------------------------------------

    /// <summary>
    /// set /slide[i]/animation[k] {trigger?, delay?, duration?}: retimes one effect in place.
    /// trigger rewrites the effect cTn's node type (and, when it crosses the click boundary, rebuilds the
    /// click-group structure so playback grouping stays consistent); delay rewrites its start condition;
    /// duration rewrites its first timed behavior. The p:timing tree stays schema-valid.
    /// </summary>
    public static string Set(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = Resolve(slidePart, address);
        var node = view.TimeNode;

        string? trigger = view.Trigger;
        long? delayMs = null;
        long? durationMs = null;

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "trigger":
                    var token = J.ScalarText(value ?? string.Empty).Trim();
                    if (!Triggers.Contains(token, StringComparer.OrdinalIgnoreCase))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"Unknown animation trigger '{token}'.",
                            "Use click (starts on click), afterPrevious or withPrevious.",
                            candidates: Triggers);
                    }

                    trigger = Triggers.First(t => string.Equals(t, token, StringComparison.OrdinalIgnoreCase));
                    break;
                case "delay":
                    delayMs = ParseSecondsMs("delay", value, max: 600, allowZero: true);
                    break;
                case "duration":
                    durationMs = ParseSecondsMs("duration", value, max: 600);
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown animation prop '{key}' for set.",
                        "Retimable animation props: trigger, delay, duration. " +
                        "Re-add the animation to change its effect/direction/color.",
                        candidates: ["trigger", "delay", "duration"]);
            }
        }

        if (durationMs is { } d)
        {
            if (string.Equals(view.Effect, "appear", StringComparison.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "Prop 'duration' does not apply to 'appear' (it is instantaneous).",
                    "Drop duration, or re-add the animation as fade for a timed entrance.");
            }

            RetimeDuration(view.Par, d);
        }

        if (delayMs is { } delay)
        {
            SetEffectDelay(node, delay);
        }

        // Trigger changes the cTn node type; when it crosses the click boundary the group layout changes,
        // so rebuild the whole mainSeq from the (unchanged) effect order with the new trigger applied.
        if (!string.Equals(trigger, view.Trigger, StringComparison.Ordinal))
        {
            node.NodeType = TriggerNodeType(trigger!);
            RebuildGroups(slidePart, applyTriggerToIndex: (view.Index, trigger!));
        }

        return address.CanonicalAnimationPath;
    }

    // ----- move (reorder) --------------------------------------------------------

    /// <summary>
    /// move /slide[i]/animation[k] before/after another animation: reorders the effect in the timeline,
    /// then rebuilds the click-group structure from the new order (each effect keeps its own trigger).
    /// Returns the moved animation's new canonical path.
    /// </summary>
    public static string Move(PresentationPart presentation, PptxAddress address, string? position)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var animations = List(slidePart);
        var from = address.AnimationIndex!.Value;
        if (from < 1 || from > animations.Count)
        {
            _ = Resolve(slidePart, address); // throws invalid_path with candidates
        }

        // Reorder by anchor reference so indices never go stale mid-shuffle.
        var pars = animations.Select(a => a.Par).ToList();
        var moving = pars[from - 1];
        var (anchorIndex, before) = ParseAnimationMoveTarget(position, address, animations);
        var anchorPar = pars[anchorIndex - 1];
        if (ReferenceEquals(anchorPar, moving))
        {
            return Units.Inv($"/slide[{address.SlideIndex}]/animation[{from}]"); // anchored to itself: a no-op
        }

        pars.Remove(moving);
        var at = pars.IndexOf(anchorPar) + (before ? 0 : 1);
        pars.Insert(at, moving);

        RebuildGroups(slidePart, orderedPars: pars);

        var newIndex = List(slidePart).FindIndex(a => ReferenceEquals(a.Par, moving)) + 1;
        return Units.Inv($"/slide[{address.SlideIndex}]/animation[{newIndex}]");
    }

    /// <summary>Parses "before /slide[i]/animation[j]" / "after ..." into (anchor index, isBefore).</summary>
    private static (int AnchorIndex, bool Before) ParseAnimationMoveTarget(
        string? position, PptxAddress address, List<AnimationView> animations)
    {
        const string usage =
            "Use \"before /slide[i]/animation[j]\" or \"after /slide[i]/animation[j]\".";
        if (string.IsNullOrWhiteSpace(position))
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, "move requires a position.", usage);
        }

        var (keyword, rest) = SplitKeyword(position.Trim());
        if (keyword is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown move position '{position}'.",
                usage + " Animations reorder relative to another animation, not by absolute index.");
        }

        var anchor = PptxAddress.Parse(rest);
        if (!anchor.IsAnimation || anchor.SlideIndex != address.SlideIndex)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The move anchor must be another animation on slide {address.SlideIndex}: '{rest}'.",
                usage);
        }

        var anchorIndex = anchor.AnimationIndex!.Value;
        if (anchorIndex < 1 || anchorIndex > animations.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Anchor animation {anchorIndex} is out of range 1..{animations.Count}."),
                usage);
        }

        return (anchorIndex, keyword == "before");
    }

    private static (string? Keyword, string Remainder) SplitKeyword(string text)
    {
        foreach (var keyword in new[] { "before", "after" })
        {
            if (text.StartsWith(keyword, StringComparison.OrdinalIgnoreCase) &&
                text.Length > keyword.Length &&
                (text[keyword.Length] == ' ' || text[keyword.Length] == ':'))
            {
                return (keyword, text[(keyword.Length + 1)..].Trim());
            }
        }

        return (null, text);
    }

    /// <summary>
    /// Rebuilds the mainSeq click-group structure from the current (or given) effect order. Each effect's
    /// own trigger decides grouping: a click effect (or the first effect) opens a new group; with/after
    /// effects join the open group. Effect pars are reused (their behaviors/ids are untouched).
    /// </summary>
    private static void RebuildGroups(
        SlidePart slidePart,
        List<P.ParallelTimeNode>? orderedPars = null,
        (int Index, string Trigger)? applyTriggerToIndex = null)
    {
        var slide = slidePart.Slide!;
        var pars = orderedPars ?? List(slidePart).Select(a => a.Par).ToList();

        var mainSequenceChildren = EnsureMainSequence(slide);

        // Detach every effect par from its current group, then drop the now-empty groups.
        foreach (var par in pars)
        {
            par.Remove();
        }

        foreach (var group in mainSequenceChildren.Elements<P.ParallelTimeNode>().ToList())
        {
            group.Remove();
        }

        var nextId = NextTimeNodeId(slide.Timing!);
        P.ChildTimeNodeList? openGroup = null;
        var index = 0;
        foreach (var par in pars)
        {
            index++;
            var effectNode = par.CommonTimeNode!;
            var trigger = applyTriggerToIndex is { } at && at.Index == index
                ? at.Trigger
                : TriggerOf(effectNode);

            if (trigger == "click" || openGroup is null)
            {
                openGroup = NewClickGroup(mainSequenceChildren, ref nextId);
            }

            openGroup.Append(par);
        }
    }

    private static string TriggerOf(P.CommonTimeNode node) => node.NodeType?.Value switch
    {
        { } t when t == P.TimeNodeValues.WithEffect => "withPrevious",
        { } t when t == P.TimeNodeValues.AfterEffect => "afterPrevious",
        _ => "click",
    };

    private static P.TimeNodeValues TriggerNodeType(string trigger) => trigger switch
    {
        "withPrevious" => P.TimeNodeValues.WithEffect,
        "afterPrevious" => P.TimeNodeValues.AfterEffect,
        _ => P.TimeNodeValues.ClickEffect,
    };

    /// <summary>Rewrites the effect's start-condition delay (in ms).</summary>
    private static void SetEffectDelay(P.CommonTimeNode node, long delayMs)
    {
        var conditions = node.GetFirstChild<P.StartConditionList>();
        if (conditions is null)
        {
            conditions = new P.StartConditionList();
            node.InsertAt(conditions, 0);
        }

        var condition = conditions.GetFirstChild<P.Condition>();
        if (condition is null)
        {
            condition = new P.Condition();
            conditions.Append(condition);
        }

        condition.Delay = delayMs.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Rewrites the first user-facing timed behavior's duration (skipping the 1ms visibility set).</summary>
    private static void RetimeDuration(P.ParallelTimeNode par, long durationMs)
    {
        foreach (var behavior in par.Descendants<P.CommonBehavior>())
        {
            if (behavior.Parent is P.SetBehavior)
            {
                continue;
            }

            if (behavior.GetFirstChild<P.CommonTimeNode>() is { } timeNode)
            {
                timeNode.Duration = durationMs.ToString(CultureInfo.InvariantCulture);
                return;
            }
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            "This animation has no timed behavior to retime.",
            "Instantaneous effects (e.g. appear) take no duration; re-add the animation as a timed effect (e.g. fade).");
    }

    private sealed record EffectSpec(
        string Class, string Effect, string Trigger, long DurationMs, long DelayMs, string Direction, string ColorHex);

    /// <summary>The preset class of an aioffice effect token ("entrance", "emphasis" or "exit").</summary>
    private static string ClassOf(string effect) =>
        EmphasisEffects.Contains(effect, StringComparer.Ordinal) ? "emphasis"
        : ExitEffects.Contains(effect, StringComparer.Ordinal) ? "exit"
        : "entrance";

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
                    "Animation props: effect, trigger, duration, delay, direction, color (colorPulse only).",
                    candidates: AddProps);
            }
        }

        var effect = Token(props, "effect") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "add animation needs the 'effect' prop.",
            "Pass an entrance (appear, fade, flyIn, wipe), emphasis (pulse, grow, spin, colorPulse) " +
            "or exit (fadeOut, flyOut, wipeOut), e.g. {\"effect\":\"fade\",\"trigger\":\"click\"}.",
            candidates: Effects);
        if (!Effects.Contains(effect, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Animation effect '{effect}' is not supported.",
                "Supported effects — entrance: appear, fade, flyIn, wipe; emphasis: pulse, grow, spin, colorPulse; " +
                "exit: fadeOut, flyOut, wipeOut. Pick the closest one and refine it in PowerPoint.",
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
            if (effect is not ("flyIn" or "wipe" or "flyOut" or "wipeOut"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop 'direction' does not apply to '{effect}'.",
                    "Only flyIn, wipe, flyOut and wipeOut take a direction (left, right, top, bottom).");
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

        var colorHex = "ED7D31"; // the accent the color pulse swings to by default
        if (props.TryGetPropertyValue("color", out var colorNode))
        {
            if (effect != "colorPulse")
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop 'color' does not apply to '{effect}'.",
                    "Only colorPulse takes a color (the fill it pulses to), e.g. {\"effect\":\"colorPulse\",\"color\":\"FF0000\"}.");
            }

            colorHex = Units.ParseColorHex("color", colorNode);
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

        return new EffectSpec(ClassOf(effect), effect, trigger, durationMs, delayMs, direction ?? "bottom", colorHex);
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

    /// <summary>The direction subtype shared by flyIn/flyOut and wipe/wipeOut presets.</summary>
    private static int DirectionSubtype(string effect, string direction) => effect switch
    {
        "flyIn" or "flyOut" => direction switch
        {
            "left" => 8,
            "right" => 2,
            "top" => 1,
            _ => 4, // bottom
        },
        _ => direction switch // wipe / wipeOut
        {
            "left" => 2,
            "right" => 8,
            "top" => 4,
            _ => 1, // bottom
        },
    };

    /// <summary>The effect par: preset cTn + per-class behaviors (visibility sets, anims, scale/rot/color).</summary>
    private static P.ParallelTimeNode BuildEffectPar(EffectSpec spec, uint shapeId, ref uint nextId)
    {
        var (presetId, presetSubtype) = spec.Effect switch
        {
            "appear" => (AppearPresetId, 0),
            "fade" => (FadePresetId, 0),
            "flyIn" or "flyOut" => (FlyInPresetId, DirectionSubtype(spec.Effect, spec.Direction)),
            "wipe" or "wipeOut" => (WipePresetId, DirectionSubtype(spec.Effect, spec.Direction)),
            "fadeOut" => (FadePresetId, 0),
            "pulse" => (PulsePresetId, 0),
            "grow" => (GrowPresetId, 0),
            "spin" => (SpinPresetId, 0),
            _ => (ColorPulsePresetId, 0), // colorPulse
        };

        var children = spec.Class == "entrance"
            ? new P.ChildTimeNodeList(BuildVisibilitySet(shapeId, "visible", delayMs: 0, ref nextId))
            : new P.ChildTimeNodeList();

        switch (spec.Effect)
        {
            case "fade":
                children.Append(BuildAnimateEffect(shapeId, "fade", transitionIn: true, spec.DurationMs, ref nextId));
                break;
            case "fadeOut":
                children.Append(BuildAnimateEffect(shapeId, "fade", transitionIn: false, spec.DurationMs, ref nextId));
                break;
            case "wipe":
            {
                var filter = spec.Direction switch
                {
                    "left" => "wipe(right)",
                    "right" => "wipe(left)",
                    "top" => "wipe(down)",
                    _ => "wipe(up)",
                };
                children.Append(BuildAnimateEffect(shapeId, filter, transitionIn: true, spec.DurationMs, ref nextId));
                break;
            }

            case "wipeOut":
            {
                // The shape disappears wiping toward the named edge.
                var filter = spec.Direction switch
                {
                    "left" => "wipe(left)",
                    "right" => "wipe(right)",
                    "top" => "wipe(up)",
                    _ => "wipe(down)",
                };
                children.Append(BuildAnimateEffect(shapeId, filter, transitionIn: false, spec.DurationMs, ref nextId));
                break;
            }

            case "flyIn":
            {
                var (fromX, fromY) = OffSlide(spec.Direction);
                children.Append(BuildAnimate(shapeId, "ppt_x", fromX, "#ppt_x", spec.DurationMs, ref nextId));
                children.Append(BuildAnimate(shapeId, "ppt_y", fromY, "#ppt_y", spec.DurationMs, ref nextId));
                break;
            }

            case "flyOut":
            {
                var (toX, toY) = OffSlide(spec.Direction);
                children.Append(BuildAnimate(shapeId, "ppt_x", "#ppt_x", toX, spec.DurationMs, ref nextId));
                children.Append(BuildAnimate(shapeId, "ppt_y", "#ppt_y", toY, spec.DurationMs, ref nextId));
                break;
            }

            case "pulse":
                children.Append(BuildAnimateScale(shapeId, PulseScale, autoReverse: true, spec.DurationMs, ref nextId));
                break;
            case "grow":
                children.Append(BuildAnimateScale(shapeId, GrowScale, autoReverse: false, spec.DurationMs, ref nextId));
                break;
            case "spin":
                children.Append(BuildAnimateRotation(shapeId, FullTurn, spec.DurationMs, ref nextId));
                break;
            case "colorPulse":
                children.Append(BuildAnimateColor(shapeId, spec.ColorHex, spec.DurationMs, ref nextId));
                break;
            default:
                break; // appear: visibility only
        }

        if (spec.Class == "exit")
        {
            // Exits hide the shape when the effect ends (PowerPoint's dur-1 convention).
            children.Append(BuildVisibilitySet(shapeId, "hidden", Math.Max(spec.DurationMs - 1, 0), ref nextId));
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
            PresetClass = spec.Class switch
            {
                "emphasis" => P.TimeNodePresetClassValues.Emphasis,
                "exit" => P.TimeNodePresetClassValues.Exit,
                _ => P.TimeNodePresetClassValues.Entrance,
            },
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

    /// <summary>The off-slide position fly effects start from (entrance) or end at (exit).</summary>
    private static (string X, string Y) OffSlide(string direction) => direction switch
    {
        "left" => ("0-#ppt_w/2", "#ppt_y"),
        "right" => ("1+#ppt_w/2", "#ppt_y"),
        "top" => ("#ppt_x", "0-#ppt_h/2"),
        _ => ("#ppt_x", "1+#ppt_h/2"), // bottom
    };

    private static P.SetBehavior BuildVisibilitySet(uint shapeId, string value, long delayMs, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode(new P.StartConditionList(new P.Condition { Delay = Inv(delayMs) }))
            {
                Id = nextId++,
                Duration = "1",
                Fill = P.TimeNodeFillValues.Hold,
            },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) }),
            new P.AttributeNameList(new P.AttributeName("style.visibility"))),
        new P.ToVariantValue(new P.StringVariantValue { Val = value }));

    /// <summary>p:animScale to the given percent*1000 (autoRev plays it back for pulse-style effects).</summary>
    private static P.AnimateScale BuildAnimateScale(uint shapeId, int scale, bool autoReverse, long durationMs, ref uint nextId)
    {
        var timeNode = new P.CommonTimeNode { Id = nextId++, Duration = Inv(durationMs), Fill = P.TimeNodeFillValues.Hold };
        if (autoReverse)
        {
            timeNode.AutoReverse = true;
        }

        return new P.AnimateScale(
            new P.CommonBehavior(
                timeNode,
                new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) })),
            new P.ToPosition { X = scale, Y = scale });
    }

    /// <summary>p:animRot by the given 60000ths-of-a-degree angle (the "r" presentation attribute).</summary>
    private static P.AnimateRotation BuildAnimateRotation(uint shapeId, int by, long durationMs, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode { Id = nextId++, Duration = Inv(durationMs), Fill = P.TimeNodeFillValues.Hold },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) }),
            new P.AttributeNameList(new P.AttributeName("r"))))
    {
        By = by,
    };

    /// <summary>p:animClr pulsing the fill color to the given hex and back (autoRev).</summary>
    private static P.AnimateColor BuildAnimateColor(uint shapeId, string hex, long durationMs, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode { Id = nextId++, Duration = Inv(durationMs), Fill = P.TimeNodeFillValues.Hold, AutoReverse = true },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) }),
            new P.AttributeNameList(new P.AttributeName("fillcolor"))),
        new P.ToColor(new A.RgbColorModelHex { Val = hex }))
    {
        ColorSpace = P.AnimateColorSpaceValues.Rgb,
    };

    private static P.AnimateEffect BuildAnimateEffect(uint shapeId, string filter, bool transitionIn, long durationMs, ref uint nextId) => new(
        new P.CommonBehavior(
            new P.CommonTimeNode { Id = nextId++, Duration = Inv(durationMs) },
            new P.TargetElement(new P.ShapeTarget { ShapeId = Inv(shapeId) })))
    {
        Transition = transitionIn ? P.AnimateEffectTransitionValues.In : P.AnimateEffectTransitionValues.Out,
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

            var presetClass = node.PresetClass?.Value;
            var cls = presetClass switch
            {
                { } c when c == P.TimeNodePresetClassValues.Emphasis => "emphasis",
                { } c when c == P.TimeNodePresetClassValues.Exit => "exit",
                _ => "entrance",
            };

            var presetId = node.PresetId?.Value;
            var effect = (cls, presetId) switch
            {
                ("entrance", AppearPresetId) => "appear",
                ("entrance", FadePresetId) => "fade",
                ("entrance", FlyInPresetId) => "flyIn",
                ("entrance", WipePresetId) => "wipe",
                ("emphasis", PulsePresetId) => "pulse",
                ("emphasis", GrowPresetId) => "grow",
                ("emphasis", SpinPresetId) => "spin",
                ("emphasis", ColorPulsePresetId) => "colorPulse",
                ("exit", FadePresetId) => "fadeOut",
                ("exit", FlyInPresetId) => "flyOut",
                ("exit", WipePresetId) => "wipeOut",
                (_, { } other) => Units.Inv($"preset{other}"), // foreign decks: truthful, not pretty
                (_, null) => "unknown",
            };

            var direction = effect switch
            {
                "flyIn" or "flyOut" => node.PresetSubtype?.Value switch
                {
                    8 => "left", 2 => "right", 1 => "top", 4 => "bottom", _ => null,
                },
                "wipe" or "wipeOut" => node.PresetSubtype?.Value switch
                {
                    2 => "left", 8 => "right", 4 => "top", 1 => "bottom", _ => null,
                },
                _ => null,
            };

            views.Add(new AnimationView(
                views.Count + 1,
                par,
                node,
                cls,
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
            Class = view.Class,
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
