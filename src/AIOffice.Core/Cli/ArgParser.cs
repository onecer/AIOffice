namespace AIOffice.Core.Cli;

/// <summary>The result of parsing an argv: verb, positionals and options.</summary>
public sealed class ParsedArgs
{
    public required string? Verb { get; init; }

    public required IReadOnlyList<string> Positionals { get; init; }

    /// <summary>Option name (without dashes) to its values, in order of appearance.</summary>
    public required IReadOnlyDictionary<string, IReadOnlyList<string>> Options { get; init; }

    /// <summary>True when the flag was present (with or without a value).</summary>
    public bool HasFlag(string name) => Options.ContainsKey(name);

    /// <summary>Last value of an option, or null. Bare flags yield "true".</summary>
    public string? GetOption(string name) =>
        Options.TryGetValue(name, out var values) && values.Count > 0 ? values[^1] : null;

    /// <summary>All values of a repeatable option.</summary>
    public IReadOnlyList<string> GetOptionValues(string name) =>
        Options.TryGetValue(name, out var values) ? values : [];
}

/// <summary>
/// Tiny argv parser, by design (no dependency): first bare token is the verb,
/// remaining bare tokens are positionals; <c>--name value</c>, <c>--name=value</c>,
/// <c>-o value</c> and bare boolean flags are options. <c>--</c> ends option
/// parsing. Values beginning with <c>@</c> can be expanded from files with
/// <see cref="ExpandAtFile"/>.
/// </summary>
public static class ArgParser
{
    /// <summary>Options that never consume the next token as a value.</summary>
    private static readonly HashSet<string> BooleanFlags = new(StringComparer.Ordinal)
    {
        "json", "pretty", "quiet", "dry-run", "track", "force",
        // find/replace sugar modifiers (M4): never eat the next token.
        "regex", "match-case", "whole-word",
        // preview mark/unmark flags.
        "tofix", "all",
    };

    public static ParsedArgs Parse(IReadOnlyList<string> argv)
    {
        string? verb = null;
        var positionals = new List<string>();
        var options = new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal);
        var optionsEnded = false;

        void AddOption(string name, string value)
        {
            if (options.TryGetValue(name, out var existing))
            {
                options[name] = [.. existing, value];
            }
            else
            {
                options[name] = [value];
            }
        }

        for (var i = 0; i < argv.Count; i++)
        {
            var token = argv[i];

            if (!optionsEnded && token == "--")
            {
                optionsEnded = true;
                continue;
            }

            if (!optionsEnded && token.Length > 1 && token[0] == '-' && !IsNegativeNumber(token))
            {
                var name = token.TrimStart('-');
                if (name.Length == 0)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Malformed option: '{token}'",
                        "Options look like --flag, --name value, --name=value or -o value.");
                }

                var equals = name.IndexOf('=');
                if (equals >= 0)
                {
                    var key = name[..equals];
                    AddOption(key, name[(equals + 1)..]);
                    continue;
                }

                // Bare boolean flag, known boolean, or option that takes the next token.
                var next = i + 1 < argv.Count ? argv[i + 1] : null;
                var nextIsValue = next is not null && (next.Length == 0 || next[0] != '-' || IsNegativeNumber(next));
                if (!BooleanFlags.Contains(name) && nextIsValue)
                {
                    AddOption(name, next!);
                    i++;
                }
                else
                {
                    AddOption(name, "true");
                }

                continue;
            }

            if (verb is null && !optionsEnded)
            {
                verb = token;
            }
            else
            {
                positionals.Add(token);
            }
        }

        return new ParsedArgs { Verb = verb, Positionals = positionals, Options = options };
    }

    /// <summary>
    /// Expands a value of the form <c>@path</c> by reading the file (through the
    /// workspace sandbox). Values not starting with <c>@</c> pass through; use
    /// <c>@@</c> to start a literal value with an at-sign.
    /// </summary>
    public static string ExpandAtFile(string value, Workspace workspace)
    {
        if (!value.StartsWith('@'))
        {
            return value;
        }

        if (value.StartsWith("@@", StringComparison.Ordinal))
        {
            return value[1..];
        }

        var path = workspace.Resolve(value[1..], mustExist: true);
        return File.ReadAllText(path);
    }

    private static bool IsNegativeNumber(string token) =>
        token.Length > 1 && token[0] == '-' && char.IsAsciiDigit(token[1]);
}
