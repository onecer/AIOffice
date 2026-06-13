using System.Globalization;
using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The shared diff pipeline behind both <c>aioffice diff</c> and the
/// <c>office_diff</c> MCP tool, so the two surfaces return a byte-identical
/// data payload. A diff is read-only: neither the current file nor the baseline
/// is mutated, and the result is always ok:true / exit 0 — the change list is
/// DATA, not an error, exactly like <c>office_validate</c>/<c>office_audit</c>.
/// <para>
/// Two baseline modes, exactly one required:
/// <list type="bullet">
/// <item>a second file — the OLD document, sandbox-resolved; its format must
///   match the current file (a mismatch is the handler's <c>invalid_args</c>);</item>
/// <item><c>--snapshot N</c> — restore snapshot N of the current file from the
///   automatic pre-edit ring into a throwaway temp file used as the baseline;
///   a missing index is <c>invalid_args</c> naming the available indices.</item>
/// </list>
/// </para>
/// </summary>
public static class DiffVerb
{
    /// <summary>The two projection levels (mirrors <see cref="DiffView"/>).</summary>
    public static IReadOnlyList<string> Views => DiffView.All;

    /// <summary>
    /// Runs a two-file diff: <paramref name="baselineResolved"/> is the OLD
    /// document (already sandbox-resolved and existence-checked by the caller),
    /// <paramref name="ctx"/>.File is the current/new one. Routes to the file's
    /// <see cref="IDiffer"/> and shapes the result envelope.
    /// </summary>
    public static Envelope RunTwoFile(
        IFormatHandler handler, CommandContext ctx, string baselineResolved, string baselineLabel, string view)
    {
        var differ = RequireDiffer(handler);
        var result = differ.Diff(ctx, baselineResolved);
        return Shape(result, baselineLabel, view);
    }

    /// <summary>
    /// Runs a snapshot diff: restores snapshot <paramref name="snapshotNumber"/>
    /// of the current file (<paramref name="ctx"/>.File) from
    /// <paramref name="snapshots"/> into a temp baseline (same extension, inside
    /// the workspace so it passes the sandbox), diffs against it, then deletes
    /// the temp. The original file and its snapshot ring are untouched.
    /// </summary>
    public static Envelope RunSnapshot(
        IFormatHandler handler, CommandContext ctx, SnapshotStore snapshots, int snapshotNumber, string view)
    {
        var differ = RequireDiffer(handler);
        var current = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "diff --snapshot needs a target file.",
            "Pass the current document, then --snapshot N to diff it against snapshot N of itself.");

        var ring = snapshots.List(current);
        if (ring.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"No snapshots exist yet for: {Path.GetFileName(current)}",
                "Snapshots are taken automatically before every successful edit; edit the file once, then diff --snapshot 1.");
        }

        var entry = ring.FirstOrDefault(e => e.Number == snapshotNumber);
        if (entry is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Snapshot {snapshotNumber} does not exist for: {Path.GetFileName(current)}",
                "Pick one of the available snapshot numbers, or run 'aioffice snapshot list <file>' to see the ring.",
                candidates: [.. ring.Select(e => e.Number.ToString(CultureInfo.InvariantCulture))]);
        }

        // Restore the snapshot bytes into a throwaway baseline alongside the file
        // (same directory + extension), so the handler resolves it like any file
        // arg and the sandbox accepts it. We copy (never Restore) so the original
        // and the ring stay untouched.
        var extension = Path.GetExtension(current);
        var baseline = Path.Combine(
            Path.GetDirectoryName(current) ?? ".",
            $".aioffice-diff-base-{Guid.NewGuid():N}{extension}");
        try
        {
            File.Copy(entry.Path, baseline, overwrite: true);
            var result = differ.Diff(ctx, baseline);
            return Shape(result, $"snapshot {snapshotNumber}", view);
        }
        finally
        {
            TryDelete(baseline);
        }
    }

    private static IDiffer RequireDiffer(IFormatHandler handler)
    {
        if (handler is IDiffer differ)
        {
            return differ;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"The {handler.Kind.ToString().ToLowerInvariant()} handler does not implement diffing in this build.",
            "Run 'aioffice doctor' for handler status; the docx/xlsx/pptx handlers all diff.");
    }

    /// <summary>
    /// Shapes the diff into the result envelope. <c>--view summary</c> trims the
    /// Before/After/Detail so the change list stays terse (path + kind only);
    /// detailed (default) keeps the full change records. Diff warnings ride into
    /// meta.warnings.
    /// </summary>
    private static Envelope Shape(DiffResult result, string baselineLabel, string view)
    {
        var summary = view.Equals(DiffView.Summary, StringComparison.Ordinal);
        var changes = summary
            ? result.Changes.Select(c => (object)new { kind = c.Kind, path = c.Path }).ToList()
            : result.Changes.Select(c => (object)c).ToList();

        var envelope = Envelope.Ok(new
        {
            changes,
            summary = result.Summary,
            baseline = baselineLabel,
            view,
        });

        return result.Warnings is { Count: > 0 } warnings
            ? envelope with { Meta = envelope.Meta with { Warnings = warnings } }
            : envelope;
    }

    /// <summary>Validates a --view value, defaulting null to detailed.</summary>
    public static string NormalizeView(string? view)
    {
        if (string.IsNullOrWhiteSpace(view))
        {
            return DiffView.Detailed;
        }

        var trimmed = view.Trim();
        if (!DiffView.All.Contains(trimmed, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown --view: '{view}'.",
                "Use 'summary' (counts + one terse line per change) or 'detailed' (full before/after; the default).",
                candidates: DiffView.All);
        }

        return trimmed;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // A stray temp baseline is harmless; it sits next to the source and
            // is overwritten/cleaned on the next diff.
        }
        catch (UnauthorizedAccessException)
        {
            // Same: best-effort cleanup.
        }
    }
}
