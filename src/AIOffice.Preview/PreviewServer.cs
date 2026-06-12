using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Preview;

/// <summary>
/// The live-preview server: a loopback-only HTTP server that renders one
/// document as interactive HTML (every addressable element carries
/// data-aio-path), tracks the user's click selection, and pushes
/// Server-Sent-Events reloads when the file changes on disk.
///
/// Routes: GET / (page), GET /events (SSE), GET+POST /selection,
/// POST /shutdown. A lockfile under <c>~/.aioffice/preview</c> advertises the
/// port to <see cref="PreviewClient"/>.
/// </summary>
public sealed class PreviewServer : IDisposable
{
    public const int PortRangeStart = 26500;
    public const int PortRangeEnd = 26600;

    private const int DebounceMs = 300;
    private const long MaxRequestBytes = 1024 * 1024;

    private readonly HttpListener _listener;
    private readonly FileSystemWatcher _watcher;
    private readonly Timer _debounce;
    private readonly SelectionStore _selection = new();
    private readonly List<SseClient> _sseClients = [];
    private readonly object _sseGate = new();
    private readonly TaskCompletionSource _shutdown = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private int _stopped;

    private PreviewServer(string filePath, Workspace workspace, HttpListener listener, int port, string lockfilePath)
    {
        FilePath = filePath;
        Workspace = workspace;
        _listener = listener;
        Port = port;
        LockfilePath = lockfilePath;

        var directory = Path.GetDirectoryName(filePath)!;
        _watcher = new FileSystemWatcher(directory, Path.GetFileName(filePath))
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size |
                           NotifyFilters.FileName | NotifyFilters.CreationTime,
        };
        _debounce = new Timer(_ => BroadcastReload(), null, Timeout.Infinite, Timeout.Infinite);
        _watcher.Changed += (_, _) => OnFileEvent();
        _watcher.Created += (_, _) => OnFileEvent();
        _watcher.Renamed += (_, _) => OnFileEvent();
        _watcher.Deleted += (_, _) => OnFileEvent();
    }

    /// <summary>Sandbox-resolved absolute path of the previewed file.</summary>
    public string FilePath { get; }

    public Workspace Workspace { get; }

    public int Port { get; }

    public string Url => string.Create(CultureInfo.InvariantCulture, $"http://127.0.0.1:{Port}/");

    public string LockfilePath { get; }

    /// <summary>Completes when the server has stopped (shutdown route, Stop or Dispose).</summary>
    public Task WaitForShutdownAsync() => _shutdown.Task;

    /// <summary>
    /// Starts a preview server for <paramref name="file"/> and returns once it
    /// is listening. <paramref name="port"/> 0 picks the first free port in
    /// 26500-26600. A stale lockfile (recorded port no longer listening) is
    /// overwritten; a live one is an error.
    /// </summary>
    public static PreviewServer Start(string file, Workspace workspace, int port = 0, string? lockDirectory = null)
    {
        var resolved = workspace.Resolve(file, mustExist: true);
        if (!File.Exists(resolved))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a file: {file}",
                "Point preview at a document file, e.g. 'aioffice preview open report.docx'.");
        }

        var extension = Path.GetExtension(resolved).ToLowerInvariant();
        if (!PreviewRenderer.SupportedExtensions.Contains(extension, StringComparer.Ordinal))
        {
            throw PreviewRenderer.UnsupportedExtension(extension);
        }

        var lockfilePath = PreviewLock.PathFor(resolved, lockDirectory);
        if (PreviewLock.TryRead(lockfilePath) is { } existing && PreviewLock.IsPortAlive(existing.Port))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"A preview for this file is already running on port {existing.Port} (pid {existing.Pid})."),
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"Open http://127.0.0.1:{existing.Port}/ in a browser, or stop it first with 'aioffice preview close {file}'."));
        }

        var (listener, boundPort) = Bind(port);
        var server = new PreviewServer(resolved, workspace, listener, boundPort, lockfilePath);
        PreviewLock.Write(lockfilePath, new PreviewLockfile(
            boundPort, Environment.ProcessId, resolved, DateTimeOffset.UtcNow));

        server._watcher.EnableRaisingEvents = true;
        _ = Task.Run(server.AcceptLoopAsync);
        return server;
    }

    /// <summary>Blocking variant for the CLI: runs until POST /shutdown (or Stop from another thread).</summary>
    public static void Open(string file, Workspace workspace, int port = 0, string? lockDirectory = null)
    {
        using var server = Start(file, workspace, port, lockDirectory);
        server.WaitForShutdownAsync().GetAwaiter().GetResult();
    }

    /// <summary>Stops the server, closes SSE streams and deletes the lockfile. Idempotent.</summary>
    public void Stop()
    {
        if (Interlocked.Exchange(ref _stopped, 1) == 1)
        {
            return;
        }

        _watcher.Dispose();
        _debounce.Dispose();

        lock (_sseGate)
        {
            foreach (var client in _sseClients)
            {
                client.Close();
            }

            _sseClients.Clear();
        }

        try
        {
            _listener.Close();
        }
        catch (ObjectDisposedException)
        {
            // Already closed.
        }

        PreviewLock.Delete(LockfilePath);
        _shutdown.TrySetResult();
    }

    public void Dispose() => Stop();

    // ----------------------------------------------------------- port binding

    private static (HttpListener Listener, int Port) Bind(int requestedPort)
    {
        if (requestedPort != 0)
        {
            return TryBind(requestedPort) is { } listener
                ? (listener, requestedPort)
                : throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    string.Create(CultureInfo.InvariantCulture, $"Port {requestedPort} is already in use."),
                    "Pass --port 0 (the default) to auto-pick a free port in 26500-26600.");
        }

        for (var candidate = PortRangeStart; candidate <= PortRangeEnd; candidate++)
        {
            if (TryBind(candidate) is { } listener)
            {
                return (listener, candidate);
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            string.Create(
                CultureInfo.InvariantCulture,
                $"No free port in the preview range {PortRangeStart}-{PortRangeEnd}."),
            "Close other previews ('aioffice preview close <file>') or pass an explicit --port.");
    }

    private static HttpListener? TryBind(int port)
    {
        var listener = new HttpListener();
        listener.Prefixes.Add(string.Create(CultureInfo.InvariantCulture, $"http://127.0.0.1:{port}/"));
        try
        {
            listener.Start();
            return listener;
        }
        catch (Exception ex) when (ex is HttpListenerException or System.Net.Sockets.SocketException)
        {
            ((IDisposable)listener).Dispose();
            return null;
        }
    }

    // ------------------------------------------------------------ http server

    private async Task AcceptLoopAsync()
    {
        while (Volatile.Read(ref _stopped) == 0)
        {
            HttpListenerContext context;
            try
            {
                context = await _listener.GetContextAsync().ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is HttpListenerException or ObjectDisposedException or InvalidOperationException)
            {
                break; // listener closed during shutdown
            }

            _ = Task.Run(() => Handle(context));
        }
    }

    private void Handle(HttpListenerContext context)
    {
        var route = (Method: context.Request.HttpMethod, Path: context.Request.Url?.AbsolutePath ?? "/");
        try
        {
            switch (route)
            {
                case ("GET", "/"):
                    HandleRoot(context);
                    break;
                case ("GET", "/events"):
                    HandleEvents(context);
                    break;
                case ("GET", "/selection"):
                    WriteJson(context.Response, 200, JsonSerializer.Serialize(SnapshotSelection(), JsonDefaults.Options));
                    break;
                case ("POST", "/selection"):
                    HandlePostSelection(context);
                    break;
                case ("POST", "/shutdown"):
                    WriteJson(context.Response, 200, Envelope.Ok(new { stopped = true }, MetaNow()).ToJson());
                    Stop();
                    break;
                default:
                    WriteJson(context.Response, 404, Envelope.Fail(
                        ErrorCodes.InvalidArgs,
                        $"No such route: {route.Method} {route.Path}",
                        "Use GET /, GET /events, GET|POST /selection or POST /shutdown.",
                        meta: MetaNow()).ToJson());
                    break;
            }
        }
        catch (AiofficeException ax)
        {
            TryWriteJson(context.Response, StatusFor(ax.Code), Envelope.FromException(ax, MetaNow()).ToJson());
        }
        catch (Exception ex)
        {
            TryWriteJson(context.Response, 500, Envelope.FromException(ex, MetaNow()).ToJson());
        }
    }

    private void HandleRoot(HttpListenerContext context)
    {
        var html = PreviewRenderer.RenderPage(FilePath, Workspace);
        WriteText(context.Response, 200, "text/html; charset=utf-8", html);
    }

    private void HandlePostSelection(HttpListenerContext context)
    {
        if (context.Request.ContentLength64 > MaxRequestBytes)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Selection body exceeds 1 MiB.",
                "Send a JSON object like {\"paths\":[\"/body/p[1]\"]}.");
        }

        string body;
        using (var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
        {
            body = reader.ReadToEnd();
        }

        JsonNode? parsed;
        try
        {
            parsed = JsonNode.Parse(body);
        }
        catch (JsonException ex)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Selection body is not valid JSON: {ex.Message}",
                "Send a JSON object like {\"paths\":[\"/body/p[1]\"]}.",
                innerException: ex);
        }

        if (parsed is not JsonObject obj || obj["paths"] is not JsonArray pathsNode)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Selection body must be an object with a 'paths' array.",
                "Send a JSON object like {\"paths\":[\"/body/p[1]\",\"/slide[2]/shape[@id=4]\"]}.");
        }

        var paths = new List<string>(pathsNode.Count);
        foreach (var node in pathsNode)
        {
            if (node is not JsonValue value || value.GetValueKind() != JsonValueKind.String)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "Every entry of 'paths' must be a string document path.",
                    "Send paths like \"/body/p[1]\" or \"/Sheet1/B2\".");
            }

            var path = value.GetValue<string>();
            PreviewPaths.Validate(path);
            paths.Add(path);
        }

        _selection.Replace(paths);
        WriteJson(context.Response, 200, JsonSerializer.Serialize(SnapshotSelection(), JsonDefaults.Options));
    }

    private void HandleEvents(HttpListenerContext context)
    {
        var response = context.Response;
        response.StatusCode = 200;
        response.ContentType = "text/event-stream";
        response.Headers["Cache-Control"] = "no-cache";
        response.SendChunked = true;

        var client = new SseClient(response);
        client.Write("retry: 1000\n\n: connected\n\n");
        lock (_sseGate)
        {
            if (Volatile.Read(ref _stopped) == 1)
            {
                client.Close();
                return;
            }

            _sseClients.Add(client);
        }

        // The response intentionally stays open; reloads are pushed by BroadcastReload.
    }

    private SelectionSnapshot SnapshotSelection()
    {
        var (paths, updatedAt) = _selection.Read();
        return new SelectionSnapshot(paths, SafeRev(), updatedAt);
    }

    private string? SafeRev()
    {
        try
        {
            return File.Exists(FilePath) ? Rev.OfFile(FilePath) : null;
        }
        catch (Exception ex) when (ex is IOException or AiofficeException)
        {
            return null; // mid-write; the next request sees the settled file
        }
    }

    private Meta MetaNow() => new() { File = FilePath, Rev = SafeRev() };

    private static int StatusFor(string code) => code switch
    {
        ErrorCodes.InvalidArgs or ErrorCodes.InvalidPath => 400,
        ErrorCodes.FileNotFound => 404,
        _ => 500,
    };

    // ----------------------------------------------------------- live reload

    private void OnFileEvent()
    {
        if (Volatile.Read(ref _stopped) == 0)
        {
            _debounce.Change(DebounceMs, Timeout.Infinite); // restart the debounce window
        }
    }

    private void BroadcastReload()
    {
        if (Volatile.Read(ref _stopped) == 1)
        {
            return;
        }

        var payload = string.Create(
            CultureInfo.InvariantCulture,
            $"event: reload\ndata: {{\"rev\":\"{SafeRev() ?? string.Empty}\"}}\n\n");

        List<SseClient> clients;
        lock (_sseGate)
        {
            clients = [.. _sseClients];
        }

        foreach (var client in clients)
        {
            if (client.TryWrite(payload))
            {
                continue;
            }

            lock (_sseGate)
            {
                _sseClients.Remove(client);
            }

            client.Close();
        }
    }

    // -------------------------------------------------------------- responses

    private static void WriteText(HttpListenerResponse response, int status, string contentType, string body)
    {
        var bytes = Encoding.UTF8.GetBytes(body);
        response.StatusCode = status;
        response.ContentType = contentType;
        response.ContentLength64 = bytes.Length;
        response.OutputStream.Write(bytes);
        response.OutputStream.Close();
    }

    private static void WriteJson(HttpListenerResponse response, int status, string json) =>
        WriteText(response, status, "application/json; charset=utf-8", json);

    private static void TryWriteJson(HttpListenerResponse response, int status, string json)
    {
        try
        {
            WriteJson(response, status, json);
        }
        catch (Exception ex) when (ex is IOException or ObjectDisposedException or InvalidOperationException or HttpListenerException)
        {
            // The client went away (or headers were already sent); nothing to salvage.
        }
    }

    /// <summary>One subscribed /events response; writes are serialized per client.</summary>
    private sealed class SseClient(HttpListenerResponse response)
    {
        private readonly object _gate = new();

        public void Write(string text)
        {
            lock (_gate)
            {
                var bytes = Encoding.UTF8.GetBytes(text);
                response.OutputStream.Write(bytes);
                response.OutputStream.Flush();
            }
        }

        public bool TryWrite(string text)
        {
            try
            {
                Write(text);
                return true;
            }
            catch (Exception ex) when (ex is IOException or ObjectDisposedException or HttpListenerException)
            {
                return false;
            }
        }

        public void Close()
        {
            try
            {
                response.OutputStream.Close();
                response.Close();
            }
            catch (Exception ex) when (ex is IOException or ObjectDisposedException or HttpListenerException)
            {
                // Already gone.
            }
        }
    }
}
