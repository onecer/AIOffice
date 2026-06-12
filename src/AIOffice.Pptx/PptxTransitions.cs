using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Slide transitions (p:transition): set via slide-level props
/// {"transition":"none|fade|push|wipe","transitionDuration":"0.5s"}, read back
/// for get and the outline view. Duration maps to the p14:dur attribute in
/// milliseconds (PowerPoint 2010+; older clients fall back to the spd default).
/// </summary>
internal static class PptxTransitions
{
    /// <summary>The transition kinds aioffice can set. Everything else is unsupported_feature.</summary>
    public static readonly IReadOnlyList<string> Kinds = ["none", "fade", "push", "wipe"];

    /// <summary>Applies transition props to a slide. Either node may be absent (flags say which were given).</summary>
    public static void Set(SlidePart slidePart, JsonNode? kindNode, bool hasKind, JsonNode? durationNode, bool hasDuration)
    {
        var slide = slidePart.Slide ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The slide part has no slide XML.",
            "The slide part is malformed; re-export the file or restore a snapshot.");

        if (hasKind)
        {
            var kind = (kindNode is null ? string.Empty : J.ScalarText(kindNode)).Trim().ToLowerInvariant();
            if (kind == "none")
            {
                if (hasDuration)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "transitionDuration cannot be combined with transition \"none\".",
                        "Drop transitionDuration, or pick a real transition: fade, push or wipe.");
                }

                slide.Transition = null;
                return;
            }

            OpenXmlElement effect = kind switch
            {
                "fade" => new P.FadeTransition(),
                "push" => new P.PushTransition(),
                "wipe" => new P.WipeTransition(),
                _ => throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Transition '{kind}' is not supported.",
                    "Supported transitions: none, fade, push, wipe. Pick the closest one and refine it in PowerPoint.",
                    candidates: Kinds),
            };
            slide.Transition = new P.Transition(effect);
        }

        if (hasDuration)
        {
            if (slide.Transition is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "transitionDuration needs a transition to apply to, but the slide has none.",
                    "Set both in one op, e.g. {\"transition\":\"fade\",\"transitionDuration\":\"0.5s\"}.");
            }

            slide.Transition.Duration = ParseDurationMs(durationNode)
                .ToString(CultureInfo.InvariantCulture);
        }
    }

    /// <summary>Duration in milliseconds. Accepts "0.5s", "500ms" or a plain number of seconds.</summary>
    private static long ParseDurationMs(JsonNode? node)
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
                var (suffix, factorToSeconds) = text switch
                {
                    _ when text.EndsWith("ms", StringComparison.Ordinal) => ("ms", 0.001),
                    _ when text.EndsWith("s", StringComparison.Ordinal) => ("s", 1.0),
                    _ => ("", 1.0),
                };
                var numberText = suffix.Length == 0 ? text : text[..^suffix.Length].Trim();
                if (double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    seconds = number * factorToSeconds;
                }
            }
        }

        if (seconds is not (> 0 and <= 600))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid transitionDuration: {node?.ToJsonString() ?? "null"}",
                "Use seconds like \"0.5s\" (or \"500ms\"); the duration must be positive and at most 600s.");
        }

        return (long)Math.Round(seconds.Value * 1000);
    }

    /// <summary>The slide's transition kind and duration ("0.5s"); null when no transition is set.</summary>
    public static (string Kind, string? Duration)? Read(SlidePart slidePart)
    {
        var transition = slidePart.Slide?.Transition;
        if (transition is null)
        {
            return null;
        }

        var kind = transition.ChildElements.FirstOrDefault() switch
        {
            null => "none",
            P.FadeTransition => "fade",
            P.PushTransition => "push",
            P.WipeTransition => "wipe",
            { } other => other.LocalName, // foreign decks: report the raw effect token truthfully
        };

        string? duration = null;
        if (long.TryParse(transition.Duration?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var ms))
        {
            duration = (ms / 1000.0).ToString("0.###", CultureInfo.InvariantCulture) + "s";
        }

        return (kind, duration);
    }
}
