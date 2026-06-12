using System.Reflection;
using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// Builds the production <see cref="HandlerRegistry"/> by discovering
/// <see cref="IFormatHandler"/> implementations in the three format assemblies.
/// <para>
/// Discovery is reflective on purpose: the format packages land in parallel,
/// and the server must keep compiling (and answering with typed
/// <c>unsupported_feature</c> envelopes) before they exist. Convention: a format
/// assembly exposes its handler as a public, non-abstract class implementing
/// <see cref="IFormatHandler"/>, constructible either with no arguments or with
/// a <see cref="SnapshotStore"/> (every other parameter must be optional). When
/// a store is injected, the handler owns pre-image snapshotting itself and the
/// kind is reported via <c>handlerManagedSnapshots</c> so the command layer
/// does not snapshot the same edit twice.
/// </para>
/// </summary>
public static class HandlerDiscovery
{
    private static readonly string[] FormatAssemblies = ["AIOffice.Word", "AIOffice.Excel", "AIOffice.Pptx"];

    /// <summary>Discovers and registers every available format handler without snapshot injection.</summary>
    public static HandlerRegistry CreateDefaultRegistry() => CreateDefaultRegistry(null, out _);

    /// <summary>
    /// Discovers and registers every available format handler. When
    /// <paramref name="snapshots"/> is provided it is injected into handler
    /// constructors that accept a <see cref="SnapshotStore"/>; those kinds are
    /// returned in <paramref name="handlerManagedSnapshots"/>.
    /// </summary>
    public static HandlerRegistry CreateDefaultRegistry(
        SnapshotStore? snapshots,
        out IReadOnlySet<DocumentKind> handlerManagedSnapshots)
    {
        var registry = new HandlerRegistry();
        var managed = new HashSet<DocumentKind>();

        foreach (var assemblyName in FormatAssemblies)
        {
            Assembly assembly;
            try
            {
                assembly = Assembly.Load(assemblyName);
            }
            catch (Exception ex) when (ex is FileNotFoundException or FileLoadException or BadImageFormatException)
            {
                continue; // Format package not present yet — registry stays honest.
            }

            foreach (var type in assembly.GetExportedTypes())
            {
                if (type.IsAbstract || !typeof(IFormatHandler).IsAssignableFrom(type))
                {
                    continue;
                }

                if (FindUsableConstructor(type) is not { } ctor)
                {
                    continue;
                }

                var (handler, injected) = Instantiate(ctor, snapshots);
                var extensions = ExtensionsFor(handler.Kind);
                if (extensions.Length > 0)
                {
                    registry.Register(handler, extensions);
                    if (injected)
                    {
                        managed.Add(handler.Kind);
                    }
                }
            }
        }

        handlerManagedSnapshots = managed;
        return registry;
    }

    /// <summary>
    /// A constructor is usable when every parameter is a SnapshotStore (we
    /// inject ours) or has a default value. Prefers the SnapshotStore overload.
    /// </summary>
    private static ConstructorInfo? FindUsableConstructor(Type type) => type
        .GetConstructors(BindingFlags.Public | BindingFlags.Instance)
        .Where(c => c.GetParameters().All(p => p.ParameterType == typeof(SnapshotStore) || p.IsOptional))
        .OrderByDescending(c => c.GetParameters().Count(p => p.ParameterType == typeof(SnapshotStore)))
        .ThenBy(c => c.GetParameters().Length)
        .FirstOrDefault();

    private static (IFormatHandler Handler, bool SnapshotsInjected) Instantiate(ConstructorInfo ctor, SnapshotStore? snapshots)
    {
        var parameters = ctor.GetParameters();
        var arguments = new object?[parameters.Length];
        var injected = false;
        for (var i = 0; i < parameters.Length; i++)
        {
            if (parameters[i].ParameterType == typeof(SnapshotStore) && snapshots is not null)
            {
                arguments[i] = snapshots;
                injected = true;
            }
            else
            {
                arguments[i] = parameters[i].IsOptional ? parameters[i].DefaultValue : null;
            }
        }

        return ((IFormatHandler)ctor.Invoke(arguments), injected);
    }

    private static string[] ExtensionsFor(DocumentKind kind) => kind switch
    {
        DocumentKind.Docx => [".docx"],
        DocumentKind.Xlsx => [".xlsx"],
        DocumentKind.Pptx => [".pptx"],
        _ => [],
    };
}
