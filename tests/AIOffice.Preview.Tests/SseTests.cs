using Xunit;

namespace AIOffice.Preview.Tests;

public sealed class SseTests : PreviewTestBase
{
    [Fact]
    public async Task Events_stream_emits_reload_after_the_file_changes_on_disk()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        // SSE responses never complete, so use a client without a global timeout.
        using var sse = new HttpClient { Timeout = Timeout.InfiniteTimeSpan };
        using var request = new HttpRequestMessage(HttpMethod.Get, server.Url + "events");
        using var response = await sse.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
        Assert.Equal("text/event-stream", response.Content.Headers.ContentType?.MediaType);

        await using var stream = await response.Content.ReadAsStreamAsync();
        using var reader = new StreamReader(stream);

        // The server greets subscribers immediately; consuming the greeting
        // proves the subscription is registered before the file is touched.
        var greeting = await reader.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(5));
        Assert.NotNull(greeting);

        var sawReload = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _ = Task.Run(async () =>
        {
            while (await reader.ReadLineAsync() is { } line)
            {
                if (line.StartsWith("event: reload", StringComparison.Ordinal))
                {
                    sawReload.TrySetResult();
                    return;
                }
            }
        });

        // Touch the file until the (300 ms debounced) reload arrives.
        var bytes = File.ReadAllBytes(file);
        for (var attempt = 0; attempt < 20 && !sawReload.Task.IsCompleted; attempt++)
        {
            File.WriteAllBytes(file, bytes);
            File.SetLastWriteTimeUtc(file, DateTime.UtcNow);
            await Task.Delay(500);
        }

        Assert.True(sawReload.Task.IsCompleted, "no reload event arrived within 10 seconds of touching the file");
    }
}
