using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// The docx semantic differ (M8): compares a baseline document against the
/// current one and reports a sorted, platform-stable change list. Body blocks
/// (paragraphs and tables) are matched with an LCS over a per-block
/// content+style hash so a matched-but-reordered block reads as <c>moved</c>
/// rather than added+removed; table cells, header/footer text and core
/// document properties are diffed in place. Output order is canonical
/// (<see cref="DiffResult.FromChanges"/> sorts by path then kind), so the same
/// two files diff identically on every platform.
/// </summary>
public sealed partial class WordHandler : IDiffer
{
    public DiffResult Diff(CommandContext ctx, string baselineFile)
    {
        var current = RequireFile(ctx, mustExist: true);

        // The baseline is the OLD document; sandbox-resolve it like any file arg.
        var baseline = ctx.Workspace.Resolve(baselineFile, mustExist: true);
        if (!File.Exists(baseline))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Baseline file not found: {baselineFile}",
                "Check the path, or pass the document you want to diff against.");
        }

        // Same-format only: a docx can only be diffed against another docx.
        var baselineExt = Path.GetExtension(baseline);
        if (!baselineExt.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot diff a .docx against a '{baselineExt}' file.",
                "Diff compares two documents of the SAME format. Pass another .docx as the baseline.");
        }

        FileSizeGuard.Ensure(baseline);

        var (newDoc, newMs, _) = OpenCopy(current, editable: false);
        var (oldDoc, oldMs, _) = OpenCopy(baseline, editable: false);
        using (newDoc)
        using (newMs)
        using (oldDoc)
        using (oldMs)
        {
            var changes = new List<DiffChange>();
            DiffBodyBlocks(oldDoc, newDoc, changes);
            DiffHeadersFooters(oldDoc, newDoc, changes);
            DiffProperties(oldDoc, newDoc, changes);
            return DiffResult.FromChanges(changes);
        }
    }

    // ------------------------------------------------------------- body blocks

    /// <summary>
    /// A top-level body block (paragraph or table) reduced to the values the
    /// diff compares: its canonical path, a text+style key used for LCS matching,
    /// and the live element for deeper (cell/style) comparison once matched.
    /// </summary>
    private sealed record Block(string Path, string Type, string MatchKey, string Text, string? Style, OpenXmlElement Element);

    private static void DiffBodyBlocks(WordprocessingDocument oldDoc, WordprocessingDocument newDoc, List<DiffChange> changes)
    {
        var oldBlocks = TopLevelBlocks(oldDoc);
        var newBlocks = TopLevelBlocks(newDoc);

        // LCS over the match key gives the longest common, in-order subsequence;
        // everything in it is "unchanged position". Blocks outside it are
        // candidates for added/removed, and we recover moves from same-key
        // blocks that fell out of the common subsequence on both sides.
        var matched = LcsMatch(oldBlocks, newBlocks);
        var matchedOld = matched.Select(m => m.OldIndex).ToHashSet();
        var matchedNew = matched.Select(m => m.NewIndex).ToHashSet();

        // Paired (matched) blocks: same key, same relative order -> compare contents.
        foreach (var (oldIndex, newIndex) in matched)
        {
            CompareMatchedBlocks(oldBlocks[oldIndex], newBlocks[newIndex], changes);
        }

        // Leftover blocks on each side. A leftover old key that also appears as a
        // leftover new key is a MOVE (same content, different position); the rest
        // are added (new-only) or removed (old-only). When a key still has
        // remaining old AND new instances we treat the overlap as modified pairs
        // (in path order) so a same-position edit is one 'modified', not add+remove.
        var leftoverOld = Enumerable.Range(0, oldBlocks.Count).Where(i => !matchedOld.Contains(i)).ToList();
        var leftoverNew = Enumerable.Range(0, newBlocks.Count).Where(i => !matchedNew.Contains(i)).ToList();

        var oldByKey = leftoverOld
            .GroupBy(i => oldBlocks[i].MatchKey, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => new Queue<int>(g), StringComparer.Ordinal);

        var consumedOld = new HashSet<int>();
        var pendingNew = new List<int>();

        foreach (var newIndex in leftoverNew)
        {
            var key = newBlocks[newIndex].MatchKey;
            if (oldByKey.TryGetValue(key, out var queue) && queue.Count > 0)
            {
                // Identical content survived but its position relative to the
                // matched anchors shifted: a move.
                var oldIndex = queue.Dequeue();
                consumedOld.Add(oldIndex);
                changes.Add(new DiffChange
                {
                    Kind = "moved",
                    Path = newBlocks[newIndex].Path,
                    Detail = "moved from " + oldBlocks[oldIndex].Path,
                });
            }
            else
            {
                pendingNew.Add(newIndex);
            }
        }

        var pendingOld = leftoverOld.Where(i => !consumedOld.Contains(i)).ToList();

        // Same-type leftovers pair up as modified (an edited block in place);
        // pairing in document order keeps it deterministic. Any surplus is a
        // genuine add or remove.
        PairModifiedOrAddRemove(oldBlocks, newBlocks, pendingOld, pendingNew, changes);
    }

    /// <summary>
    /// Pairs leftover old/new blocks of the same type (in path order) as
    /// <c>modified</c>, then reports the unpaired remainder as <c>removed</c>
    /// (old-only) or <c>added</c> (new-only).
    /// </summary>
    private static void PairModifiedOrAddRemove(
        IReadOnlyList<Block> oldBlocks, IReadOnlyList<Block> newBlocks,
        List<int> pendingOld, List<int> pendingNew, List<DiffChange> changes)
    {
        var oldByType = pendingOld
            .GroupBy(i => oldBlocks[i].Type, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => new Queue<int>(g.OrderBy(i => i)), StringComparer.Ordinal);

        var newByType = pendingNew
            .GroupBy(i => newBlocks[i].Type, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => new Queue<int>(g.OrderBy(i => i)), StringComparer.Ordinal);

        foreach (var type in oldByType.Keys.Concat(newByType.Keys).Distinct(StringComparer.Ordinal).ToList())
        {
            var olds = oldByType.TryGetValue(type, out var oq) ? oq : new Queue<int>();
            var news = newByType.TryGetValue(type, out var nq) ? nq : new Queue<int>();

            while (olds.Count > 0 && news.Count > 0)
            {
                CompareMatchedBlocks(oldBlocks[olds.Dequeue()], newBlocks[news.Dequeue()], changes, forceModified: true);
            }

            while (olds.Count > 0)
            {
                var block = oldBlocks[olds.Dequeue()];
                changes.Add(new DiffChange { Kind = "removed", Path = block.Path, Detail = BlockDetail(block) });
            }

            while (news.Count > 0)
            {
                var block = newBlocks[news.Dequeue()];
                changes.Add(new DiffChange { Kind = "added", Path = block.Path, Detail = BlockDetail(block) });
            }
        }
    }

    /// <summary>
    /// Two blocks the differ paired up: emit a concise <c>modified</c> for a text
    /// or style change (paragraphs) and recurse into matched table cells. When
    /// <paramref name="forceModified"/> is set the keys differ by construction, so
    /// always summarise the change even if the cheap text/style probes coincide.
    /// </summary>
    private static void CompareMatchedBlocks(Block oldBlock, Block newBlock, List<DiffChange> changes, bool forceModified = false)
    {
        if (oldBlock.Type == "p" && newBlock.Type == "p")
        {
            var textChanged = !string.Equals(oldBlock.Text, newBlock.Text, StringComparison.Ordinal);
            var styleChanged = !string.Equals(NormalizeStyle(oldBlock.Style), NormalizeStyle(newBlock.Style), StringComparison.Ordinal);

            if (styleChanged)
            {
                changes.Add(new DiffChange
                {
                    Kind = "modified",
                    Path = newBlock.Path,
                    Before = NormalizeStyle(oldBlock.Style),
                    After = NormalizeStyle(newBlock.Style),
                    Detail = "style",
                });
            }

            if (textChanged || (forceModified && !styleChanged))
            {
                changes.Add(new DiffChange
                {
                    Kind = "modified",
                    Path = newBlock.Path,
                    Before = oldBlock.Text,
                    After = newBlock.Text,
                    Detail = "text",
                });
            }

            return;
        }

        if (oldBlock.Type == "table" && newBlock.Type == "table")
        {
            DiffTableCells((Table)oldBlock.Element, (Table)newBlock.Element, newBlock.Path, changes);
            return;
        }

        // Differing block types in the same slot: the old one went, the new arrived.
        changes.Add(new DiffChange { Kind = "removed", Path = oldBlock.Path, Detail = BlockDetail(oldBlock) });
        changes.Add(new DiffChange { Kind = "added", Path = newBlock.Path, Detail = BlockDetail(newBlock) });
    }

    /// <summary>Cell-by-cell text diff of two matched tables; reports at the current cell path.</summary>
    private static void DiffTableCells(Table oldTable, Table newTable, string newTablePath, List<DiffChange> changes)
    {
        var oldRows = oldTable.Elements<TableRow>().ToList();
        var newRows = newTable.Elements<TableRow>().ToList();
        var rowCount = Math.Max(oldRows.Count, newRows.Count);

        for (var r = 0; r < rowCount; r++)
        {
            var oldCells = r < oldRows.Count ? oldRows[r].Elements<TableCell>().ToList() : [];
            var newCells = r < newRows.Count ? newRows[r].Elements<TableCell>().ToList() : [];
            var cellCount = Math.Max(oldCells.Count, newCells.Count);

            for (var c = 0; c < cellCount; c++)
            {
                var cellPath = string.Create(
                    CultureInfo.InvariantCulture, $"{newTablePath}/tr[{r + 1}]/tc[{c + 1}]");

                var oldText = c < oldCells.Count ? CellText(oldCells[c]) : null;
                var newText = c < newCells.Count ? CellText(newCells[c]) : null;

                if (oldText is null && newText is not null)
                {
                    changes.Add(new DiffChange { Kind = "added", Path = cellPath, Detail = "cell" });
                }
                else if (oldText is not null && newText is null)
                {
                    // A cell present in the baseline but gone now: report at the
                    // positional path (the table matched, so the index aligns).
                    changes.Add(new DiffChange { Kind = "removed", Path = cellPath, Detail = "cell" });
                }
                else if (oldText is not null && newText is not null &&
                         !string.Equals(oldText, newText, StringComparison.Ordinal))
                {
                    changes.Add(new DiffChange
                    {
                        Kind = "modified",
                        Path = cellPath,
                        Before = oldText,
                        After = newText,
                        Detail = "cell",
                    });
                }
            }
        }
    }

    // ------------------------------------------------------- headers / footers

    /// <summary>
    /// Header/footer paragraph text changes, addressed at /header[i]/p[j] etc.
    /// Roots are paired by part order (the same order Resolve and EnumerateAll
    /// use); a root that exists on only one side reports its paragraphs as
    /// added/removed.
    /// </summary>
    private static void DiffHeadersFooters(WordprocessingDocument oldDoc, WordprocessingDocument newDoc, List<DiffChange> changes)
    {
        DiffRootGroup(
            WordAddress.HeaderFooterRoots(oldDoc).Where(r => r.Type == "header").ToList(),
            WordAddress.HeaderFooterRoots(newDoc).Where(r => r.Type == "header").ToList(),
            changes);
        DiffRootGroup(
            WordAddress.HeaderFooterRoots(oldDoc).Where(r => r.Type == "footer").ToList(),
            WordAddress.HeaderFooterRoots(newDoc).Where(r => r.Type == "footer").ToList(),
            changes);
    }

    private static void DiffRootGroup(List<ResolvedNode> oldRoots, List<ResolvedNode> newRoots, List<DiffChange> changes)
    {
        var count = Math.Max(oldRoots.Count, newRoots.Count);
        for (var i = 0; i < count; i++)
        {
            var oldRoot = i < oldRoots.Count ? oldRoots[i] : null;
            var newRoot = i < newRoots.Count ? newRoots[i] : null;

            var oldParas = oldRoot is null
                ? []
                : oldRoot.Element.Elements<Paragraph>().ToList();
            var newParas = newRoot is null
                ? []
                : newRoot.Element.Elements<Paragraph>().ToList();
            var rootPath = (newRoot ?? oldRoot!).CanonicalPath;
            var paraCount = Math.Max(oldParas.Count, newParas.Count);

            for (var p = 0; p < paraCount; p++)
            {
                var paraPath = string.Create(CultureInfo.InvariantCulture, $"{rootPath}/p[{p + 1}]");
                var oldText = p < oldParas.Count ? oldParas[p].InnerText : null;
                var newText = p < newParas.Count ? newParas[p].InnerText : null;

                if (oldText is null && newText is not null)
                {
                    changes.Add(new DiffChange { Kind = "added", Path = paraPath, Detail = "headerFooter" });
                }
                else if (oldText is not null && newText is null)
                {
                    changes.Add(new DiffChange { Kind = "removed", Path = paraPath, Detail = "headerFooter" });
                }
                else if (oldText is not null && newText is not null &&
                         !string.Equals(oldText, newText, StringComparison.Ordinal))
                {
                    changes.Add(new DiffChange
                    {
                        Kind = "modified",
                        Path = paraPath,
                        Before = oldText,
                        After = newText,
                        Detail = "headerFooter",
                    });
                }
            }
        }
    }

    // -------------------------------------------------------------- properties

    /// <summary>Core document-property changes, reported at /properties with the property name as Detail.</summary>
    private static void DiffProperties(WordprocessingDocument oldDoc, WordprocessingDocument newDoc, List<DiffChange> changes)
    {
        var oldCore = ReadCoreProps(oldDoc);
        var newCore = ReadCoreProps(newDoc);

        CompareProperty(changes, "title", oldCore.Title, newCore.Title);
        CompareProperty(changes, "author", oldCore.Creator, newCore.Creator);
        CompareProperty(changes, "subject", oldCore.Subject, newCore.Subject);
        CompareProperty(changes, "keywords", oldCore.Keywords, newCore.Keywords);
        CompareProperty(changes, "category", oldCore.Category, newCore.Category);
        CompareProperty(changes, "comments", oldCore.Description, newCore.Description);
    }

    private static void CompareProperty(List<DiffChange> changes, string name, string? oldValue, string? newValue)
    {
        var before = oldValue is { Length: > 0 } ? oldValue : null;
        var after = newValue is { Length: > 0 } ? newValue : null;
        if (!string.Equals(before, after, StringComparison.Ordinal))
        {
            changes.Add(new DiffChange
            {
                Kind = "modified",
                Path = "/properties",
                Before = before,
                After = after,
                Detail = name,
            });
        }
    }

    // ------------------------------------------------------------------ blocks

    /// <summary>Top-level body blocks (paragraphs + tables) with their canonical paths.</summary>
    private static List<Block> TopLevelBlocks(WordprocessingDocument doc)
    {
        var blocks = new List<Block>();
        if (doc.MainDocumentPart?.Document?.Body is not { } body)
        {
            return blocks;
        }

        var pIndex = 0;
        var tableIndex = 0;
        foreach (var child in body.ChildElements)
        {
            switch (child)
            {
                case Paragraph paragraph:
                {
                    pIndex++;
                    var path = string.Create(CultureInfo.InvariantCulture, $"/body/p[{pIndex}]");
                    var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    var text = paragraph.InnerText;
                    blocks.Add(new Block(
                        path, "p",
                        MatchKey: "p" + NormalizeStyle(style) + "" + text,
                        Text: text, Style: style, Element: paragraph));
                    break;
                }

                case Table table:
                {
                    tableIndex++;
                    var path = string.Create(CultureInfo.InvariantCulture, $"/body/table[{tableIndex}]");
                    var text = table.InnerText;
                    blocks.Add(new Block(
                        path, "table",
                        MatchKey: "table" + text,
                        Text: text, Style: null, Element: table));
                    break;
                }

                default:
                    break; // sectPr and other plumbing are not addressable blocks
            }
        }

        return blocks;
    }

    private static string CellText(TableCell cell) => cell.InnerText;

    private static string BlockDetail(Block block) => block.Type == "table"
        ? "table"
        : NormalizeStyle(block.Style) == "Normal"
            ? "paragraph"
            : NormalizeStyle(block.Style);

    /// <summary>A null/empty paragraph style is Word's default "Normal".</summary>
    private static string NormalizeStyle(string? style) => style is { Length: > 0 } ? style : "Normal";

    // ----------------------------------------------------------------- LCS

    /// <summary>
    /// The longest common subsequence of two block lists by match key, returned
    /// as old/new index pairs in order. Standard DP LCS; ties resolve toward the
    /// earlier old index so the matching is deterministic.
    /// </summary>
    private static List<(int OldIndex, int NewIndex)> LcsMatch(IReadOnlyList<Block> olds, IReadOnlyList<Block> news)
    {
        var m = olds.Count;
        var n = news.Count;
        var dp = new int[m + 1, n + 1];

        for (var i = m - 1; i >= 0; i--)
        {
            for (var j = n - 1; j >= 0; j--)
            {
                dp[i, j] = string.Equals(olds[i].MatchKey, news[j].MatchKey, StringComparison.Ordinal)
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);
            }
        }

        var pairs = new List<(int, int)>();
        int x = 0, y = 0;
        while (x < m && y < n)
        {
            if (string.Equals(olds[x].MatchKey, news[y].MatchKey, StringComparison.Ordinal))
            {
                pairs.Add((x, y));
                x++;
                y++;
            }
            else if (dp[x + 1, y] >= dp[x, y + 1])
            {
                x++;
            }
            else
            {
                y++;
            }
        }

        return pairs;
    }
}
