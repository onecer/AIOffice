using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The shared command layer behind both the MCP tools and the CLI verbs: one
/// method per capability, every method returning a fully-stamped
/// <see cref="Envelope"/> and never throwing. Cross-cutting mechanics live
/// here exactly once — sandbox resolution, <c>expect_rev</c> optimistic
/// locking (checked BEFORE any write), automatic pre-image snapshots,
/// rev/elapsedMs stamping — while format-specific work is delegated to the
/// <see cref="IFormatHandler"/> registered for the file's extension.
/// <para>
/// Handler contract for mutating verbs: before <see cref="IFormatHandler.Edit"/>
/// or an in-place <see cref="IFormatHandler.Template"/> runs, the service
/// injects <c>args["snapshot"]</c> (the pre-image snapshot number, absent on
/// dry-run) and <c>args["dryRun"]</c> (normalized bool) into
/// <see cref="CommandContext.Args"/>; handlers echo them in their data payload.
/// </para>
/// </summary>
public sealed class CommandService
{
    private static readonly IReadOnlyList<string> KindNames = ["docx", "xlsx", "pptx"];

    private readonly SnapshotStore _snapshots;
    private readonly string _snapshotDir;
    private readonly IReadOnlySet<DocumentKind> _handlerManagedSnapshots;

    /// <param name="workspace">Sandbox every file argument is resolved against.</param>
    /// <param name="handlers">Format handler registry (see <see cref="HandlerDiscovery"/>).</param>
    /// <param name="snapshotBaseDir">Snapshot ring location; defaults to <c>~/.aioffice/snapshots</c>.</param>
    /// <param name="handlerManagedSnapshots">Kinds whose handler received a <see cref="SnapshotStore"/> and snapshots pre-images itself; the service then skips its own pre-image snapshot to keep the ring duplicate-free.</param>
    public CommandService(
        Workspace workspace,
        HandlerRegistry handlers,
        string? snapshotBaseDir = null,
        IReadOnlySet<DocumentKind>? handlerManagedSnapshots = null)
    {
        Workspace = workspace;
        Handlers = handlers;
        _snapshotDir = snapshotBaseDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".aioffice", "snapshots");
        _snapshots = new SnapshotStore(_snapshotDir);
        _handlerManagedSnapshots = handlerManagedSnapshots ?? new HashSet<DocumentKind>();
    }

    public Workspace Workspace { get; }

    public HandlerRegistry Handlers { get; }

    // ── format verbs ────────────────────────────────────────────────────────

    public Envelope Create(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass a target path ending in .docx, .xlsx or .pptx.");
        var resolved = Workspace.Resolve(file);
        if (File.Exists(resolved) && !OptionalBool(a, "overwrite", false))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Target already exists: {file}",
                "Pass overwrite:true to replace it, or choose a different path.");
        }

        var kind = OptionalString(a, "kind");
        var handler = kind is null ? Handlers.Resolve(resolved) : ResolveByKind(kind);
        if (Path.GetDirectoryName(resolved) is { Length: > 0 } parent)
        {
            Directory.CreateDirectory(parent);
        }

        // M5 markdown/csv bridge (same routing as the CLI's create --from).
        return OptionalString(a, "from") is { } from
            ? Bridge.CreateFrom(handler, Context(resolved, a), from)
            : handler.Create(Context(resolved, a));
    });

    public Envelope Read(JsonObject args) => FormatVerb(args, static (h, ctx) =>
    {
        // M5: markdown/csv are single-format bridge views — asking the wrong
        // format reports unsupported_feature with the views it does have.
        Bridge.GuardBridgeView(h.Kind, OptionalString(ctx.Args, "view"));
        return h.Read(ctx);
    });

    public Envelope Query(JsonObject args) => FormatVerb(args, static (h, ctx) => h.Query(ctx));

    public Envelope Get(JsonObject args) => FormatVerb(args, static (h, ctx) => h.Get(ctx));

    public Envelope Render(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass the document path (inside the workspace).");
        var resolved = Workspace.Resolve(file, mustExist: true);
        var handler = Handlers.Resolve(resolved);

        // png/pdf are cross-format plumbing (handler artifact -> headless
        // browser); everything else goes straight to the handler.
        var to = OptionalString(a, "to");
        if (to is "png" or "pdf")
        {
            if (OptionalString(a, "output") is { } output)
            {
                a["output"] = Workspace.Resolve(output); // the render verbs expect a resolved path
            }

            return to == "png"
                ? AIOffice.Render.PngRenderVerb.Execute(handler, Context(resolved, a))
                : AIOffice.Render.PdfRenderVerb.Execute(handler, Context(resolved, a));
        }

        return handler.Render(Context(resolved, a));
    });

    public Envelope Validate(JsonObject args) => FormatVerb(args, static (h, ctx) => h.Validate(ctx));

    public Envelope Audit(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass the document to audit (.docx/.xlsx/.pptx).");
        var resolved = Workspace.Resolve(file, mustExist: true);
        var handler = Handlers.Resolve(resolved);
        if (handler is not IAuditor auditor)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"The {handler.Kind.ToString().ToLowerInvariant()} handler does not implement auditing in this build.",
                "Use office_status for handler health; the docx/xlsx/pptx handlers all audit.");
        }

        var opts = ParseAuditOptions(a);
        a["category"] = opts.Category;
        a["minSeverity"] = opts.MinSeverity;
        a["fix"] = opts.Fix;
        return AuditVerb.Run(auditor, Context(resolved, a), opts);
    });

    public Envelope Diff(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass the current document to diff.");
        var resolved = Workspace.Resolve(file, mustExist: true);
        var handler = Handlers.Resolve(resolved);
        var view = DiffVerb.NormalizeView(OptionalString(a, "view"));

        // Exactly one baseline: another file (other) OR a snapshot index.
        var other = OptionalString(a, "other");
        var snapshot = OptionalInt(a, "snapshot");

        if (other is null && snapshot is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "diff needs a baseline: pass 'other' (another same-format file) or 'snapshot' (a snapshot number).",
                "Two files: office_diff {file, other}. Against a snapshot: office_diff {file, snapshot:1}.");
        }

        if (other is not null && snapshot is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "diff takes EITHER 'other' OR 'snapshot', not both.",
                "Drop one: 'other' diffs against that file; 'snapshot' diffs against the file's own snapshot ring.");
        }

        var ctx = Context(resolved, a);

        if (snapshot is { } n)
        {
            if (n <= 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'snapshot' must be a positive snapshot number, got {n}.",
                    "Call file_snapshot {file, action:\"list\"} to see the available snapshot numbers.");
            }

            return DiffVerb.RunSnapshot(handler, ctx, _snapshots, n, view);
        }

        var baseline = Workspace.Resolve(other!, mustExist: true);
        return DiffVerb.RunTwoFile(handler, ctx, baseline, other!, view);
    });

    public Envelope Edit(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass the document to edit.");
        var resolved = Workspace.Resolve(file, mustExist: true);
        var handler = Handlers.Resolve(resolved);

        var opsNode = a["ops"] ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "'ops' is required.",
            "Pass a JSON array like [{\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"Hi\"}}].");
        var ops = EditOp.ParseBatch(opsNode.ToJsonString(JsonDefaults.Options));

        // M3 cross-doc bridge: pptx chart ops may pull categories/series from a
        // workbook ({"dataFrom":"book.xlsx!Sheet1/A1:B5"}); expanded BEFORE the
        // rev guard and snapshot so a bad source range writes nothing.
        ops = CrossDocDataFrom.Expand(ops, handler.Kind, Workspace, Handlers);

        // M4 find/replace sugar: a root-scoped replace op ("/") fans out over
        // the format's default scopes (docx body+headers+footers, every sheet,
        // every slide incl. notes); results are aggregated after the edit.
        ops = ReplaceSugar.ExpandDocumentScopes(
            ops, handler.Kind, resolved, OptionalBool(a, "track", false), out var replaceExpansion);

        GuardRev(resolved, OptionalString(a, "expect_rev"));

        // M2 attribution: op props.author > tool arg author > AIOFFICE_AUTHOR > handler default.
        if (OptionalString(a, "author") is null &&
            Environment.GetEnvironmentVariable("AIOFFICE_AUTHOR") is { Length: > 0 } envAuthor)
        {
            a["author"] = envAuthor;
        }

        var dryRun = OptionalBool(a, "dry_run", false);
        a["dryRun"] = dryRun;
        var envelope = WithPreImageSnapshot(resolved,
            takeSnapshot: !dryRun && !_handlerManagedSnapshots.Contains(handler.Kind), a,
            () => handler.Edit(Context(resolved, a), ops));
        return replaceExpansion is null
            ? envelope
            : ReplaceSugar.Aggregate(envelope, replaceExpansion);
    });

    public Envelope Template(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass the template document containing {{key}} placeholders.");
        var resolved = Workspace.Resolve(file, mustExist: true);
        var handler = Handlers.Resolve(resolved);

        if (a["data"] is not JsonObject)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "'data' must be a JSON object of string values.",
                "Pass a merge map like {\"client\":\"ACME Corp\",\"date\":\"2026-06-12\"}.");
        }

        var output = OptionalString(a, "output");
        var inPlace = output is null;
        if (output is not null)
        {
            var outResolved = Workspace.Resolve(output);
            if (File.Exists(outResolved) && !OptionalBool(a, "overwrite", false))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Output already exists: {output}",
                    "Pass overwrite:true to replace it, or choose a different output path.");
            }
        }

        GuardRev(resolved, OptionalString(a, "expect_rev"));
        return WithPreImageSnapshot(resolved,
            takeSnapshot: inPlace && !_handlerManagedSnapshots.Contains(handler.Kind), a,
            () => handler.Template(Context(resolved, a)));
    });

    public Envelope Convert(JsonObject args) => Run(args, a =>
    {
        var src = RequireString(a, "src", "Pass the source document, e.g. report.docx.");
        var dest = RequireString(a, "dest", "Pass the destination, e.g. deck.pptx (a fresh file is created).");
        return ConvertVerb.Run(Workspace, Handlers, ConvertibleHandlersByKind(), src, dest);
    });

    /// <summary>
    /// Resolves the office handlers by kind for the neutral-model bridge, reading
    /// them out of the extension registry. Kinds whose package has not shipped are
    /// simply absent — <see cref="ConvertVerb"/> reports an honest
    /// <c>unsupported_feature</c> if one is needed.
    /// </summary>
    private IReadOnlyDictionary<DocumentKind, IFormatHandler> ConvertibleHandlersByKind()
    {
        var byKind = new Dictionary<DocumentKind, IFormatHandler>();
        foreach (var (ext, kind) in new[]
        {
            (".docx", DocumentKind.Docx), (".xlsx", DocumentKind.Xlsx), (".pptx", DocumentKind.Pptx),
        })
        {
            try
            {
                byKind[kind] = Handlers.Resolve(ext);
            }
            catch (AiofficeException)
            {
                // Package not present; leave the kind unmapped.
            }
        }

        return byKind;
    }

    // ── format-agnostic verbs ───────────────────────────────────────────────

    public Envelope Snapshot(JsonObject args) => Run(args, a =>
    {
        var file = RequireString(a, "file", "Pass the document whose snapshot ring you want.");
        var resolved = Workspace.Resolve(file);
        var action = OptionalString(a, "action") ?? "list";

        switch (action)
        {
            case "list":
                var ring = _snapshots.List(resolved);
                return Envelope.Ok(new
                {
                    snapshots = ring.Select(e => new
                    {
                        n = e.Number,
                        at = e.CreatedUtc,
                        rev = e.Rev,
                        bytes = e.SizeBytes,
                        trigger = "auto",
                    }).ToArray(),
                });

            case "restore":
                GuardRev(resolved, OptionalString(a, "expect_rev"));
                var restored = _snapshots.Restore(resolved, OptionalInt(a, "n"));
                return Envelope.Ok(new { restored = restored.Number, rev = Rev.OfFile(resolved) });

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown snapshot action: '{action}'.",
                    "Use action:\"list\" to see the ring or action:\"restore\" (with optional n) to roll back.",
                    candidates: ["list", "restore"]);
        }
    });

    public Envelope Status() => Run([], _ =>
    {
        var checks = new List<object>();
        var healthy = true;

        bool Check(string name, Func<string> probe)
        {
            try
            {
                checks.Add(new { name, ok = true, detail = probe() });
                return true;
            }
            catch (Exception ex)
            {
                checks.Add(new { name, ok = false, detail = ex.Message });
                return false;
            }
        }

        healthy &= Check("workspace_writable", () =>
        {
            var probe = Path.Combine(Workspace.Root, ".aioffice-doctor-" + Guid.NewGuid().ToString("N"));
            File.WriteAllText(probe, "ok");
            File.Delete(probe);
            return Workspace.Root;
        });

        healthy &= Check("snapshot_store_writable", () =>
        {
            Directory.CreateDirectory(_snapshotDir);
            var probe = Path.Combine(_snapshotDir, ".aioffice-doctor-" + Guid.NewGuid().ToString("N"));
            File.WriteAllText(probe, "ok");
            File.Delete(probe);
            return _snapshotDir;
        });

        healthy &= Check("format_handlers", () =>
        {
            var extensions = Handlers.KnownExtensions.Order(StringComparer.Ordinal).ToArray();
            return extensions.Length > 0
                ? string.Join(", ", extensions)
                : throw new InvalidOperationException("no format handlers registered yet (M0 in progress)");
        });

        long count = 0, bytes = 0;
        if (Directory.Exists(_snapshotDir))
        {
            foreach (var f in Directory.EnumerateFiles(_snapshotDir, "*.snap", SearchOption.AllDirectories))
            {
                count++;
                bytes += new FileInfo(f).Length;
            }
        }

        return Envelope.Ok(new
        {
            healthy,
            version = Meta.ToolVersion,
            runtime = new
            {
                dotnet = Environment.Version.ToString(),
                os = OperatingSystem.IsMacOS() ? "macos" : OperatingSystem.IsWindows() ? "windows" : "linux",
                arch = RuntimeInformation.OSArchitecture.ToString().ToLowerInvariant(),
            },
            workspace = Workspace.Root,
            snapshotStore = new { path = _snapshotDir, count, bytes },
            capabilities = new
            {
                mcpTools = ToolCatalog.Names.Count,
                mcpToolNames = ToolCatalog.Names,
                verbs = SurfaceSchema.VerbNames.Count,
                formats = new[] { "docx", "xlsx", "pptx" },
                convert = new
                {
                    sources = ConvertVerb.SupportedSources,
                    contentTargets = ConvertVerb.SupportedContentTargets,
                    renderTargets = ConvertVerb.SupportedRenderTargets,
                },
                renderTargets = new[] { "html", "svg", "text", "png", "pdf" },
                auditCategories = AuditOptions.Categories,
            },
            checks,
        });
    });

    public Envelope PreviewOpen(JsonObject args) => Run(args, a => PreviewTools.Open(Workspace, a));

    public Envelope PreviewSelection(JsonObject args) => Run(args, a => PreviewTools.Selection(Workspace, a));

    public Envelope Help(JsonObject args) => Run(args, a => Envelope.Ok(HelpTopics.Get(OptionalString(a, "topic"))));

    public Envelope Schema(JsonObject args) => Run(args, a => Envelope.Ok(SurfaceSchema.Build(OptionalString(a, "verb"))));

    // ── cross-cutting plumbing ──────────────────────────────────────────────

    private Envelope FormatVerb(JsonObject args, Func<IFormatHandler, CommandContext, Envelope> invoke) =>
        Run(args, a =>
        {
            var file = RequireString(a, "file", "Pass the document path (inside the workspace).");
            var resolved = Workspace.Resolve(file, mustExist: true);
            return invoke(Handlers.Resolve(resolved), Context(resolved, a));
        });

    private CommandContext Context(string resolvedFile, JsonObject args) =>
        new() { Workspace = Workspace, File = resolvedFile, Args = args };

    /// <summary>
    /// Times the body, converts any exception into a failure envelope, and
    /// stamps meta (file, rev-after-call, elapsedMs). Handler-provided meta
    /// values win; only missing ones are filled in.
    /// </summary>
    private Envelope Run(JsonObject args, Func<JsonObject, Envelope> body)
    {
        var stopwatch = Stopwatch.StartNew();
        string? userFile = null;
        if (args["file"] is JsonValue v && v.TryGetValue<string>(out var s))
        {
            userFile = s;
        }

        Envelope envelope;
        try
        {
            envelope = body(args);
        }
        catch (Exception ex)
        {
            envelope = Envelope.FromException(ex);
        }

        var meta = envelope.Meta;
        var file = meta.File ?? userFile;
        var rev = meta.Rev;
        if (rev is null && file is not null)
        {
            try
            {
                var resolved = Workspace.Resolve(file);
                if (File.Exists(resolved))
                {
                    rev = Rev.OfFile(resolved);
                }
            }
            catch (AiofficeException)
            {
                // Sandbox-denied or unresolvable: the envelope already reports it.
            }
        }

        return envelope with { Meta = meta with { File = file, Rev = rev, ElapsedMs = stopwatch.ElapsedMilliseconds } };
    }

    /// <summary>Throws <c>stale_address</c> when the file's current rev differs from <paramref name="expectRev"/>.</summary>
    private static void GuardRev(string resolvedFile, string? expectRev)
    {
        if (expectRev is null)
        {
            return;
        }

        var current = Rev.OfFile(resolvedFile);
        if (!current.Equals(expectRev, StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.StaleAddress,
                $"expect_rev mismatch: the file is at rev {current}, you expected {expectRev}.",
                "The file changed since you last read it. Re-run office_read/office_get to refresh paths and rev, then retry; nothing was written.");
        }
    }

    /// <summary>
    /// Snapshots the pre-image (when <paramref name="takeSnapshot"/>), injects the
    /// snapshot number into <paramref name="args"/>, runs the mutation, and discards
    /// the snapshot again when the mutation did not succeed — failed batches must
    /// not pollute the undo ring.
    /// </summary>
    private Envelope WithPreImageSnapshot(string resolvedFile, bool takeSnapshot, JsonObject args, Func<Envelope> mutate)
    {
        SnapshotEntry? pre = null;
        if (takeSnapshot)
        {
            pre = _snapshots.Save(resolvedFile);
            args["snapshot"] = pre.Number;
        }

        Envelope envelope;
        try
        {
            envelope = mutate();
        }
        catch
        {
            Discard(pre);
            throw;
        }

        if (!envelope.IsOk)
        {
            Discard(pre);
        }

        return envelope;

        static void Discard(SnapshotEntry? entry)
        {
            if (entry is null)
            {
                return;
            }

            try
            {
                File.Delete(entry.Path);
            }
            catch (IOException)
            {
                // Best effort: a stray pre-image snapshot is harmless.
            }
        }
    }

    /// <summary>Parses category/severity/fix args into an <see cref="AuditOptions"/>, validating the enums.</summary>
    private static AuditOptions ParseAuditOptions(JsonObject args)
    {
        var category = OptionalString(args, "category") ?? "all";
        if (!AuditOptions.Categories.Contains(category, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown category: '{category}'.",
                "Use one of: accessibility, quality, all (default).",
                candidates: AuditOptions.Categories);
        }

        var severity = OptionalString(args, "severity") ?? "info";
        if (!AuditOptions.Severities.Contains(severity, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown severity: '{severity}'.",
                "Use one of: error, warning, info (the minimum level to report; default info).",
                candidates: AuditOptions.Severities);
        }

        return new AuditOptions
        {
            Category = category,
            MinSeverity = severity,
            Fix = OptionalBool(args, "fix", false),
        };
    }

    private IFormatHandler ResolveByKind(string kind)
    {
        if (!KindNames.Contains(kind, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown kind: '{kind}'.",
                "Use one of: docx, xlsx, pptx.",
                candidates: KindNames);
        }

        return Handlers.Resolve("kind-override." + kind.ToLowerInvariant());
    }

    // ── argument extraction (typed invalid_args instead of crashes) ─────────

    private static string RequireString(JsonObject args, string name, string suggestion) =>
        OptionalString(args, name) ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs, $"'{name}' is required.", suggestion);

    private static string? OptionalString(JsonObject args, string name)
    {
        var node = args[name];
        if (node is null)
        {
            return null;
        }

        if (node is JsonValue value && value.TryGetValue<string>(out var s))
        {
            return s;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' must be a string.",
            $"Pass {name} as a JSON string.");
    }

    private static bool OptionalBool(JsonObject args, string name, bool fallback)
    {
        var node = args[name];
        if (node is null)
        {
            return fallback;
        }

        if (node is JsonValue value && value.TryGetValue<bool>(out var b))
        {
            return b;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' must be a boolean.",
            $"Pass {name} as true or false.");
    }

    private static int? OptionalInt(JsonObject args, string name)
    {
        var node = args[name];
        if (node is null)
        {
            return null;
        }

        if (node is JsonValue value && value.TryGetValue<int>(out var i))
        {
            return i;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' must be an integer.",
            $"Pass {name} as a JSON number.");
    }
}
