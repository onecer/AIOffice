using System.Diagnostics;
using System.Globalization;
using AIOffice.Core;
using AIOffice.Core.Cli;
using AIOffice.Preview;

namespace AIOffice.Cli;

/// <summary>
/// The <c>aioffice preview</c> verb: <c>open</c> runs a blocking live-preview
/// server (the startup envelope with url/port/pid is printed and flushed
/// BEFORE blocking so parents can scrape it), <c>selection</c> reads the paths
/// the user clicked in the browser, <c>close</c> shuts the server down.
/// </summary>
internal static class PreviewVerbs
{
    private static readonly string[] Actions = ["open", "selection", "close", "mark", "unmark", "marks", "goto"];

    /// <summary>Dispatches one preview action; returns the process exit code.</summary>
    public static int Run(ParsedArgs args, Workspace workspace, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        try
        {
            var action = Positional(args, 0, "action");
            var file = Positional(args, 1, "file");

            return action switch
            {
                "open" => Open(args, workspace, file, stopwatch, print),
                "selection" => Selection(workspace, file, stopwatch, print),
                "close" => Close(workspace, file, stopwatch, print),
                "mark" => Mark(args, workspace, file, stopwatch, print),
                "unmark" => Unmark(args, workspace, file, stopwatch, print),
                "marks" => Marks(workspace, file, stopwatch, print),
                "goto" => Goto(args, workspace, file, stopwatch, print),
                _ => throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown preview action: '{action}'.",
                    "Use open/selection/close/mark/unmark/marks/goto, e.g. 'aioffice preview mark report.docx /body/p[3] --note overflows'.",
                    candidates: Actions),
            };
        }
        catch (Exception ex)
        {
            return print(Envelope.FromException(ex, new Meta { ElapsedMs = stopwatch.ElapsedMilliseconds }));
        }
    }

    private static int Open(ParsedArgs args, Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var server = PreviewServer.Start(file, workspace, ParsePort(args));
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true; // stop cleanly so the lockfile is deleted
            server.Stop();
        };

        var envelope = Envelope.Ok(
            new
            {
                url = server.Url,
                port = server.Port,
                pid = Environment.ProcessId,
                lockfile = server.LockfilePath,
            },
            MetaFor(server.FilePath, stopwatch));

        // Print + flush BEFORE blocking: parents (MCP preview_open, scripts)
        // scrape url/port from stdout while the server keeps running.
        var exitCode = print(envelope);
        Console.Out.Flush();

        server.WaitForShutdownAsync().GetAwaiter().GetResult();
        return exitCode;
    }

    private static int Selection(Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var resolved = workspace.Resolve(file);
        var snapshot = PreviewClient.GetSelection(resolved);
        return print(Envelope.Ok(
            new { paths = snapshot.Paths, rev = snapshot.Rev, updatedAt = snapshot.UpdatedAt },
            MetaFor(resolved, stopwatch)));
    }

    private static int Close(Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var resolved = workspace.Resolve(file);
        PreviewClient.Close(resolved);
        return print(Envelope.Ok(new { closed = true }, MetaFor(resolved, stopwatch)));
    }

    private static int Mark(ParsedArgs args, Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var path = Positional(args, 2, "path");
        var resolved = workspace.Resolve(file);
        var snapshot = PreviewClient.AddMark(
            resolved, path, args.GetOption("color"), args.GetOption("note"), args.GetOption("find"), args.HasFlag("tofix"));
        return print(Envelope.Ok(MarksData(snapshot), MetaFor(resolved, stopwatch)));
    }

    private static int Unmark(ParsedArgs args, Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var resolved = workspace.Resolve(file);
        var path = args.Positionals.Count > 2 ? args.Positionals[2] : null;
        var snapshot = args.HasFlag("all")
            ? PreviewClient.ClearMarks(resolved)
            : PreviewClient.RemoveMark(
                resolved,
                path ?? throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "Pass a <path> or --all.",
                    "e.g. 'aioffice preview unmark report.docx /body/p[3]', or '--all' to clear every mark."));
        return print(Envelope.Ok(MarksData(snapshot), MetaFor(resolved, stopwatch)));
    }

    private static int Marks(Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var resolved = workspace.Resolve(file);
        return print(Envelope.Ok(MarksData(PreviewClient.GetMarks(resolved)), MetaFor(resolved, stopwatch)));
    }

    private static int Goto(ParsedArgs args, Workspace workspace, string file, Stopwatch stopwatch, Func<Envelope, int> print)
    {
        var path = Positional(args, 2, "path");
        var resolved = workspace.Resolve(file);
        PreviewClient.Goto(resolved, path);
        return print(Envelope.Ok(new { scrolledTo = path }, MetaFor(resolved, stopwatch)));
    }

    private static object MarksData(MarksSnapshot snapshot) =>
        new { marks = snapshot.Marks, rev = snapshot.Rev, updatedAt = snapshot.UpdatedAt };

    private static int ParsePort(ParsedArgs args)
    {
        var text = args.GetOption("port");
        if (text is null)
        {
            return 0;
        }

        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var port) ||
            port is < 0 or > 65535)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"--port must be 0-65535, got '{text}'.",
                "Omit --port to auto-pick a free port in 26500-26600, or pass a valid TCP port.");
        }

        return port;
    }

    private static string Positional(ParsedArgs args, int index, string name)
    {
        if (args.Positionals.Count <= index)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Missing required argument: <{name}>.",
                "Usage: aioffice preview <open|selection|close|mark|unmark|marks|goto> <file> [<path>] [--port N].",
                candidates: index == 0 ? Actions : null);
        }

        return args.Positionals[index];
    }

    private static Meta MetaFor(string resolvedFile, Stopwatch stopwatch)
    {
        string? rev = null;
        try
        {
            rev = File.Exists(resolvedFile) ? Rev.OfFile(resolvedFile) : null;
        }
        catch (Exception ex) when (ex is IOException or AiofficeException)
        {
            // Meta is best-effort; the data payload already answered.
        }

        return new Meta { File = resolvedFile, Rev = rev, ElapsedMs = stopwatch.ElapsedMilliseconds };
    }
}
