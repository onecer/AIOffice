using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Pptx;

/// <summary>EMU/centimeter/point conversions and prop-value parsing for pptx geometry.</summary>
internal static class Units
{
    public const long EmuPerCm = 360_000;
    public const double EmuPerPixel = 9_525; // 96 dpi

    public static long CmToEmu(double cm) => (long)Math.Round(cm * EmuPerCm);

    public static double EmuToCm(long emu) => Math.Round((double)emu / EmuPerCm, 2);

    public static double EmuToPx(long emu) => emu / EmuPerPixel;

    /// <summary>Invariant-culture interpolation shorthand.</summary>
    public static string Inv(FormattableString text) => FormattableString.Invariant(text);

    /// <summary>
    /// Parses a length prop into EMU. Plain numbers are centimeters; strings
    /// accept cm/emu/pt/in/px suffixes ("5cm", "1800000emu").
    /// </summary>
    public static long ParseLengthEmu(string key, JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (TryNumber(value, out var cm))
            {
                return CmToEmu(cm);
            }

            if (value.TryGetValue<string>(out var raw))
            {
                var text = raw.Trim().ToLowerInvariant();
                var (suffix, factorToEmu) = text switch
                {
                    _ when text.EndsWith("emu", StringComparison.Ordinal) => ("emu", 1.0),
                    _ when text.EndsWith("cm", StringComparison.Ordinal) => ("cm", EmuPerCm),
                    _ when text.EndsWith("pt", StringComparison.Ordinal) => ("pt", 12_700.0),
                    _ when text.EndsWith("in", StringComparison.Ordinal) => ("in", 914_400.0),
                    _ when text.EndsWith("px", StringComparison.Ordinal) => ("px", EmuPerPixel),
                    _ => ("", EmuPerCm),
                };

                var numberText = suffix.Length == 0 ? text : text[..^suffix.Length].Trim();
                if (double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    return (long)Math.Round(number * factorToEmu);
                }
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a valid length: {node?.ToJsonString() ?? "null"}",
            "Use centimeters as a number (e.g. 5 or 2.5) or a suffixed string like \"5cm\", \"1800000emu\", \"36pt\", \"1in\".");
    }

    /// <summary>Parses an RRGGBB color (leading '#' allowed) into uppercase hex.</summary>
    public static string ParseColorHex(string key, JsonNode? node)
    {
        if (node is JsonValue value && value.TryGetValue<string>(out var raw))
        {
            var text = raw.Trim().TrimStart('#').ToUpperInvariant();
            if (text.Length == 6 && text.All(Uri.IsHexDigit))
            {
                return text;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a valid color: {node?.ToJsonString() ?? "null"}",
            "Use a 6-digit RRGGBB hex string like \"FF0000\" or \"#3366CC\".");
    }

    /// <summary>Parses a font size in points into OOXML hundredths of a point.</summary>
    public static int ParseFontSizeHundredths(string key, JsonNode? node)
    {
        if (node is JsonValue value && TryNumber(value, out var points) && points is >= 1 and <= 4000)
        {
            return (int)Math.Round(points * 100);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a valid font size: {node?.ToJsonString() ?? "null"}",
            "Use a point size between 1 and 4000, e.g. 18 or 24.5.");
    }

    /// <summary>
    /// Extracts a number from a JsonValue regardless of its backing type
    /// (parsed JSON numbers and in-memory int/long/double/decimal values alike).
    /// </summary>
    public static bool TryNumber(JsonValue value, out double number)
    {
        if (value.TryGetValue(out number))
        {
            return true;
        }

        if (value.TryGetValue<int>(out var i))
        {
            number = i;
            return true;
        }

        if (value.TryGetValue<long>(out var l))
        {
            number = l;
            return true;
        }

        if (value.TryGetValue<float>(out var f))
        {
            number = f;
            return true;
        }

        if (value.TryGetValue<decimal>(out var m))
        {
            number = (double)m;
            return true;
        }

        number = 0;
        return false;
    }
}

/// <summary>Tolerant getters over the verb-args <see cref="JsonObject"/>.</summary>
internal static class J
{
    public static string? Str(JsonObject args, string key)
    {
        if (!args.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        return node is JsonValue value && value.TryGetValue<string>(out var text) ? text : node.ToJsonString();
    }

    public static bool? Bool(JsonObject args, string key)
    {
        if (args.TryGetPropertyValue(key, out var node) && node is JsonValue value)
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

        return null;
    }

    /// <summary>A scalar JsonNode rendered as a plain string (numbers/bools invariant).</summary>
    public static string ScalarText(JsonNode node) =>
        node is JsonValue value && value.TryGetValue<string>(out var text) ? text : node.ToJsonString();
}
