using System.Reflection;
using AIOffice.Core;

namespace AIOffice.Cli;

/// <summary>Doctor-facing status of one format assembly.</summary>
public sealed record HandlerStatus(
    string Assembly,
    string Kind,
    IReadOnlyList<string> Extensions,
    string Status,
    string? HandlerType,
    string? SnapshotOwner);

/// <summary>Everything Discover() learns about the available handlers.</summary>
public sealed record DiscoveredHandlers(
    HandlerRegistry Registry,
    IReadOnlyDictionary<DocumentKind, IFormatHandler> ByKind,
    IReadOnlyList<HandlerStatus> Statuses,
    /// <summary>Kinds whose handler received a SnapshotStore and snapshots pre-images itself.</summary>
    IReadOnlySet<DocumentKind> HandlerManagedSnapshots);

/// <summary>
/// Late-bound discovery of <see cref="IFormatHandler"/> implementations in the
/// sibling format assemblies (AIOffice.Word/Excel/Pptx). The contract for a
/// format package: ship one public, non-abstract implementation of
/// <c>AIOffice.Core.IFormatHandler</c> constructible either with no arguments
/// or with a <see cref="SnapshotStore"/> (which is injected, and then the
/// handler owns pre-edit snapshotting). Until a package ships one, its verbs
/// answer with a typed <c>unsupported_feature</c> envelope instead of crashing.
/// </summary>
public static class HandlerDiscovery
{
    private static readonly (string Assembly, DocumentKind Kind, string[] Extensions)[] FormatAssemblies =
    [
        ("AIOffice.Word", DocumentKind.Docx, [".docx"]),
        ("AIOffice.Excel", DocumentKind.Xlsx, [".xlsx"]),
        ("AIOffice.Pptx", DocumentKind.Pptx, [".pptx"]),
    ];

    /// <summary>Discovers handlers, registers them, and reports per-assembly status.</summary>
    public static DiscoveredHandlers Discover(SnapshotStore snapshots)
    {
        var registry = new HandlerRegistry();
        var byKind = new Dictionary<DocumentKind, IFormatHandler>();
        var statuses = new List<HandlerStatus>();
        var handlerManaged = new HashSet<DocumentKind>();

        foreach (var (assemblyName, kind, extensions) in FormatAssemblies)
        {
            var kindName = kind.ToString().ToLowerInvariant();
            IFormatHandler? handler = null;
            string? handlerType = null;
            var snapshotsInjected = false;
            string status;

            try
            {
                var assembly = Assembly.Load(new AssemblyName(assemblyName));
                var implementation = assembly
                    .GetExportedTypes()
                    .FirstOrDefault(t =>
                        !t.IsAbstract &&
                        typeof(IFormatHandler).IsAssignableFrom(t) &&
                        FindUsableConstructor(t) is not null);

                if (implementation is null)
                {
                    status = "pending (no IFormatHandler implementation yet)";
                }
                else
                {
                    (handler, snapshotsInjected) = Instantiate(FindUsableConstructor(implementation)!, snapshots);
                    handlerType = implementation.FullName;
                    status = "ready";
                }
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                status = $"unavailable ({ex.GetType().Name}: {ex.Message})";
            }

            if (handler is not null)
            {
                registry.Register(handler, extensions);
                byKind[kind] = handler;
                if (snapshotsInjected)
                {
                    handlerManaged.Add(kind);
                }
            }

            statuses.Add(new HandlerStatus(
                assemblyName, kindName, extensions, status, handlerType,
                handler is null ? null : snapshotsInjected ? "handler" : "cli"));
        }

        return new DiscoveredHandlers(registry, byKind, statuses, handlerManaged);
    }

    /// <summary>
    /// A constructor is usable when every parameter is a SnapshotStore (we
    /// inject ours) or has a default value. Prefers the SnapshotStore overload.
    /// </summary>
    private static ConstructorInfo? FindUsableConstructor(Type type)
    {
        return type
            .GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .Where(c => c.GetParameters().All(p => p.ParameterType == typeof(SnapshotStore) || p.IsOptional))
            .OrderByDescending(c => c.GetParameters().Count(p => p.ParameterType == typeof(SnapshotStore)))
            .ThenBy(c => c.GetParameters().Length)
            .FirstOrDefault();
    }

    private static (IFormatHandler Handler, bool SnapshotsInjected) Instantiate(ConstructorInfo ctor, SnapshotStore snapshots)
    {
        var parameters = ctor.GetParameters();
        var arguments = new object?[parameters.Length];
        var injected = false;
        for (var i = 0; i < parameters.Length; i++)
        {
            if (parameters[i].ParameterType == typeof(SnapshotStore))
            {
                arguments[i] = snapshots;
                injected = true;
            }
            else
            {
                arguments[i] = parameters[i].DefaultValue;
            }
        }

        return ((IFormatHandler)ctor.Invoke(arguments), injected);
    }

    /// <summary>
    /// Lazily resolves the MCP server entry point: a public static
    /// <c>RunAsync</c> method in the AIOffice.Mcp assembly returning a Task,
    /// whose parameters are drawn from: <c>string</c>/<c>string?</c> (the
    /// workspace root), <c>string[]</c>/<c>IReadOnlyList&lt;string&gt;</c>
    /// (raw argv) and <c>CancellationToken</c>. The richest satisfiable
    /// overload wins. Returns null while the MCP package has not shipped yet.
    /// </summary>
    public static Func<string?, string[], CancellationToken, Task<int>>? FindMcpEntryPoint(out string status)
    {
        try
        {
            var assembly = Assembly.Load(new AssemblyName("AIOffice.Mcp"));
            var method = assembly
                .GetExportedTypes()
                .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
                .Where(m =>
                    m.Name == "RunAsync" &&
                    typeof(Task).IsAssignableFrom(m.ReturnType) &&
                    m.GetParameters().All(p => IsMappableParameter(p.ParameterType)))
                .OrderByDescending(m => m.GetParameters().Length)
                .FirstOrDefault();

            if (method is null)
            {
                status = "pending (no public static RunAsync yet)";
                return null;
            }

            status = $"ready ({method.DeclaringType!.FullName}.RunAsync)";
            return BuildInvoker(method);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            status = $"unavailable ({ex.GetType().Name}: {ex.Message})";
            return null;
        }
    }

    private static bool IsMappableParameter(Type type) =>
        type == typeof(string) || type == typeof(string[]) ||
        type == typeof(IReadOnlyList<string>) || type == typeof(CancellationToken);

    private static Func<string?, string[], CancellationToken, Task<int>> BuildInvoker(MethodInfo method)
    {
        return async (workspaceRoot, args, cancellationToken) =>
        {
            var parameters = method.GetParameters();
            var arguments = new object?[parameters.Length];
            for (var i = 0; i < parameters.Length; i++)
            {
                var type = parameters[i].ParameterType;
                arguments[i] =
                    type == typeof(string) ? workspaceRoot :
                    type == typeof(CancellationToken) ? (object)cancellationToken :
                    args; // string[] / IReadOnlyList<string>
            }

            var task = (Task)method.Invoke(null, arguments)!;
            await task.ConfigureAwait(false);
            return task is Task<int> withExit ? withExit.Result : ExitCodes.Ok;
        };
    }
}
