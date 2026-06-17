using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using AIOffice.Core;
using AIOffice.Core.Cli;

namespace AIOffice.Cli;

/// <summary>
/// The verbs implemented natively by the CLI (no format handler involved):
/// snapshot, doctor, schema, help and version.
/// </summary>
public sealed class NativeVerbs
{
    private static readonly string[] HelpTopics =
    [
        "addressing", "selectors", "properties-docx", "properties-xlsx", "properties-pptx", "errors",
        "equations", "embeds", "rtl", "sections", "audit", "diff", "convert",
        // v1.2 additive topics.
        "smartart", "connectors", "number-formats", "structural-fields",
        // v1.3 additive topics.
        "chart-polish", "conditional-format", "themes", "3d-models", "form-fields", "animations",
        // v1.4 additive topics.
        "formulas", "data-tables", "mail-merge", "page-borders", "zoom", "table-styles",
        // v1.5 additive topics.
        "scenarios", "goal-seek", "table-formulas", "building-blocks", "embedded-fonts",
        "action-buttons", "layouts", "line-numbers",
        // v1.7 additive topics.
        "print-setup", "masters",
        // v1.9 additive topics.
        "render-engines",
    ];

    private readonly Workspace _workspace;
    private readonly SnapshotStore _snapshots;
    private readonly IReadOnlyList<HandlerStatus> _handlerStatuses;

    public NativeVerbs(Workspace workspace, SnapshotStore snapshots, IReadOnlyList<HandlerStatus> handlerStatuses)
    {
        _workspace = workspace;
        _snapshots = snapshots;
        _handlerStatuses = handlerStatuses;
    }

    // ----- snapshot --------------------------------------------------------

    public Envelope Snapshot(ParsedArgs args)
    {
        if (args.Positionals.Count < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "snapshot needs an action and a file: aioffice snapshot <list|restore> <file> [n].",
                "Example: aioffice snapshot list report.docx — then aioffice snapshot restore report.docx 3.",
                candidates: ["list", "restore"]);
        }

        var action = args.Positionals[0];
        var file = _workspace.Resolve(args.Positionals[1]);

        switch (action)
        {
            case "list":
            {
                var ring = _snapshots.List(file);
                return Envelope.Ok(new
                {
                    file,
                    capacity = SnapshotStore.Capacity,
                    count = ring.Count,
                    snapshots = ring.Select(s => new
                    {
                        number = s.Number,
                        sizeBytes = s.SizeBytes,
                        createdUtc = s.CreatedUtc,
                        rev = s.Rev,
                    }),
                });
            }

            case "restore":
            {
                int? number = null;
                if (args.Positionals.Count > 2)
                {
                    if (!int.TryParse(args.Positionals[2], NumberStyles.None, CultureInfo.InvariantCulture, out var n))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"Snapshot number must be a positive integer, got '{args.Positionals[2]}'.",
                            "Run 'aioffice snapshot list <file>' to see available numbers.");
                    }

                    number = n;
                }

                var restored = _snapshots.Restore(file, number);
                return Envelope.Ok(new
                {
                    file,
                    restored = new { number = restored.Number, rev = restored.Rev, createdUtc = restored.CreatedUtc },
                });
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown snapshot action: '{action}'.",
                    "Use 'aioffice snapshot list <file>' or 'aioffice snapshot restore <file> [n]'.",
                    candidates: ["list", "restore"]);
        }
    }

    // ----- doctor ----------------------------------------------------------

    public Envelope Doctor()
    {
        _ = HandlerDiscovery.FindMcpEntryPoint(out var mcpStatus);

        return Envelope.Ok(new
        {
            name = "aioffice",
            version = Meta.ToolVersion,
            runtime = new
            {
                framework = RuntimeInformation.FrameworkDescription,
                os = RuntimeInformation.OSDescription,
                architecture = RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
            },
            workspace = new
            {
                root = _workspace.Root,
                writable = ProbeWritable(_workspace.Root),
            },
            limits = new
            {
                // M3: unlimited by default; AIOFFICE_MAX_FILE_MB is an opt-in cap.
                maxFileMb = FileSizeGuard.MaxFileMb is { } mb ? (object)mb : "unlimited",
                maxFileMbDefault = "unlimited",
                maxFileMbEnv = FileSizeGuard.EnvVar,
            },
            handlers = _handlerStatuses.Select(h => new
            {
                kind = h.Kind,
                assembly = h.Assembly,
                extensions = h.Extensions,
                status = h.Status,
                type = h.HandlerType,
                snapshotOwner = h.SnapshotOwner,
            }),
            mcp = new { assembly = "AIOffice.Mcp", status = mcpStatus },
            browser = AIOffice.Render.BrowserLocator.Probe(),
            renderers = Renderers(),
            preview = PreviewLockDirInfo(),
            dependencies = new[]
            {
                DependencyInfo("DocumentFormat.OpenXml"),
                DependencyInfo("ClosedXML"),
                DependencyInfo("ModelContextProtocol", "ModelContextProtocol.Core"),
            },
            snapshots = new
            {
                directory = Path.GetDirectoryName(_snapshots.RingDirectory(Path.Combine(_workspace.Root, "example.docx"))),
                capacityPerFile = SnapshotStore.Capacity,
            },
            capabilities = Capabilities(),
        });
    }

    /// <summary>
    /// A one-call introspection of the whole aioffice surface so an agent can
    /// learn what it can do without probing each verb: the verb + MCP tool
    /// counts, the formats it handles, the convert source/target matrix, the
    /// render targets and the audit categories.
    /// </summary>
    private static object Capabilities() => new
    {
        surfaceVersion = Meta.SurfaceVersion,
        verbs = CommandSurface.VerbNames.Count,
        verbNames = CommandSurface.VerbNames,
        mcpTools = AIOffice.Mcp.ToolCatalog.Names.Count,
        mcpToolNames = AIOffice.Mcp.ToolCatalog.Names,
        formats = new[] { "docx", "xlsx", "pptx" },
        convert = new
        {
            sources = AIOffice.Mcp.ConvertVerb.SupportedSources,
            contentTargets = AIOffice.Mcp.ConvertVerb.SupportedContentTargets,
            renderTargets = AIOffice.Mcp.ConvertVerb.SupportedRenderTargets,
        },
        renderTargets = new[] { "html", "svg", "text", "png", "pdf" },
        auditCategories = AuditOptions.Categories,
    };

    /// <summary>
    /// The render engines available on this machine (v1.9). <c>chromium</c> is
    /// the default engine and screenshots/prints aioffice's own HTML/SVG
    /// projection; <c>libreoffice</c> (soffice) is the optional TRUE-fidelity
    /// engine that hands the original document to LibreOffice, and <c>poppler</c>
    /// (pdftoppm) rasterizes a PDF page to PNG for it. The existing top-level
    /// <c>browser</c> field is unchanged; this is an additive summary.
    /// </summary>
    private static object Renderers()
    {
        var browser = AIOffice.Render.BrowserLocator.Probe();
        var soffice = AIOffice.Render.SofficeLocator.Probe();
        return new
        {
            chromium = new { engine = "chromium", found = browser.Found, path = browser.Path, kind = browser.Kind },
            libreoffice = new { engine = "soffice", found = soffice.Found, path = soffice.Path },
            poppler = new { tool = "pdftoppm", found = soffice.Pdftoppm, path = soffice.PdftoppmPath },
        };
    }

    /// <summary>Status of the preview lockfile directory (where running servers advertise their port).</summary>
    private static object PreviewLockDirInfo()
    {
        var directory = AIOffice.Preview.PreviewLock.DefaultDirectory;
        var lockfiles = 0;
        try
        {
            if (Directory.Exists(directory))
            {
                lockfiles = Directory.EnumerateFiles(directory, "*.json").Count();
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Diagnostics stay best-effort; the directory is recreated on demand.
        }

        return new { lockDirectory = directory, exists = Directory.Exists(directory), lockfiles };
    }

    private static object DependencyInfo(string name, params string[] fallbackAssemblies)
    {
        foreach (var assemblyName in (string[])[name, .. fallbackAssemblies])
        {
            try
            {
                var assembly = Assembly.Load(new AssemblyName(assemblyName));
                var version =
                    assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                    ?? assembly.GetName().Version?.ToString()
                    ?? "unknown";
                var plus = version.IndexOf('+', StringComparison.Ordinal);
                return new
                {
                    name,
                    assembly = assemblyName,
                    version = plus > 0 ? version[..plus] : version,
                    loaded = true,
                };
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // try the next candidate assembly name
            }
        }

        return new { name, assembly = (string?)null, version = (string?)null, loaded = false };
    }

    private static bool ProbeWritable(string directory)
    {
        try
        {
            var probe = Path.Combine(directory, ".aioffice-doctor-" + Guid.NewGuid().ToString("N") + ".tmp");
            File.WriteAllText(probe, "probe");
            File.Delete(probe);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    // ----- schema ----------------------------------------------------------

    public static Envelope Schema(ParsedArgs args)
    {
        if (args.Positionals.Count > 0)
        {
            var name = args.Positionals[0];
            var verb = CommandSurface.Find(name) ?? throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown verb: '{name}'.",
                "Run 'aioffice schema' for the whole surface.",
                candidates: CommandSurface.NearestVerbs(name));
            return Envelope.Ok(new { verb });
        }

        return Envelope.Ok(new
        {
            name = "aioffice",
            version = Meta.ToolVersion,
            surfaceVersion = Meta.SurfaceVersion,
            envelope = new
            {
                description = "Every command prints exactly one JSON object to stdout: " +
                              "{ok, data|null, error{code,message,suggestion,candidates?}|null, " +
                              "meta{file?,rev?,elapsedMs,version,warnings?}}.",
                rev = "First 12 hex chars of the SHA-256 of the file bytes; pass it back via edit --expect-rev.",
            },
            globalFlags = CommandSurface.GlobalFlags,
            verbs = CommandSurface.Verbs,
            addressing = GrammarPointers.Addressing,
            selectors = GrammarPointers.Selectors,
            errorCodes = ErrorCodes.All.Select(code => new
            {
                code,
                exitCode = ExitCodes.ForErrorCode(code),
                // formula_not_evaluated / find_no_match are warning-level: they ride
                // in meta.warnings and the command still exits 0. They appear in the
                // frozen error-code table (CONTRACT §2) but never set an exit code.
                warningLevel = code is ErrorCodes.FormulaNotEvaluated or ErrorCodes.FindNoMatch,
            }),
            warningCodes = WarningCodes.All,
            exitCodes = new
            {
                ok = ExitCodes.Ok,
                userError = ExitCodes.UserError,
                internalError = ExitCodes.InternalError,
                sandboxDenied = ExitCodes.SandboxDenied,
                unsupportedFeature = ExitCodes.UnsupportedFeature,
            },
        });
    }

    // ----- help ------------------------------------------------------------

    public static Envelope Help(ParsedArgs args)
    {
        if (args.Positionals.Count == 0)
        {
            return Envelope.Ok(new
            {
                topic = "overview",
                topics = HelpTopics,
                text = LoadTopic("overview"),
                verbs = CommandSurface.Verbs.Select(v => new { v.Name, v.Summary, v.Usage }),
            });
        }

        var topic = args.Positionals[0];
        if (!HelpTopics.Contains(topic, StringComparer.Ordinal))
        {
            // Verb names are valid help topics too: forward to the schema entry.
            if (CommandSurface.Find(topic) is { } verb)
            {
                return Envelope.Ok(new { topic, verb });
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown help topic: '{topic}'.",
                "Run 'aioffice help' to list topics, or 'aioffice schema' for the machine-readable surface.",
                candidates: HelpTopics);
        }

        return Envelope.Ok(new { topic, text = LoadTopic(topic) });
    }

    private static string LoadTopic(string topic)
    {
        var assembly = typeof(NativeVerbs).Assembly;
        var suffix = $".{topic}.md";
        var resourceName = assembly
            .GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith(suffix, StringComparison.Ordinal));
        if (resourceName is null)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"Embedded help topic missing: {topic}.",
                "This is a packaging bug in aioffice; run 'aioffice schema' for the surface meanwhile.");
        }

        using var stream = assembly.GetManifestResourceStream(resourceName)!;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    // ----- version ---------------------------------------------------------

    public static Envelope Version() => Envelope.Ok(new
    {
        name = "aioffice",
        version = Meta.ToolVersion,
        runtime = RuntimeInformation.FrameworkDescription,
    });
}
