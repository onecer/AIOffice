using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>Restores the process-wide probe cache to a clean probe on dispose.</summary>
public sealed class ProbeCacheReset : IDisposable
{
    public void Dispose() => BrowserLocator.Probe(refresh: true);
}

public sealed class BrowserLocatorTests : IDisposable
{
    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    [Fact]
    public void Env_var_with_existing_file_wins_the_probe()
    {
        var stub = _tmp.WriteStubBrowser("fake-chrome", "exit 0");

        using var reset = new ProbeCacheReset();
        using var env = new EnvVarScope(BrowserLocator.EnvVar, stub);
        var info = BrowserLocator.Probe(refresh: true);

        Assert.True(info.Found);
        Assert.Equal(stub, info.Path);
        Assert.Equal("chrome", info.Kind);
    }

    [Theory]
    [InlineData("fake-msedge", "edge")]
    [InlineData("fake-chromium", "chromium")]
    [InlineData("my-browser", "custom")]
    public void Kind_is_classified_from_the_binary_name(string name, string expectedKind)
    {
        var stub = _tmp.WriteStubBrowser(name, "exit 0");

        using var reset = new ProbeCacheReset();
        using var env = new EnvVarScope(BrowserLocator.EnvVar, stub);
        var info = BrowserLocator.Probe(refresh: true);

        Assert.True(info.Found);
        Assert.Equal(expectedKind, info.Kind);
    }

    [Fact]
    public void Env_var_pointing_at_a_missing_file_falls_through_to_other_probes()
    {
        var bogus = _tmp.PathOf("does-not-exist");

        using var reset = new ProbeCacheReset();
        using var env = new EnvVarScope(BrowserLocator.EnvVar, bogus);
        var info = BrowserLocator.Probe(refresh: true);

        // Whatever the machine has (or nothing), the bogus path must not win.
        Assert.NotEqual(bogus, info.Path);
    }

    [Fact]
    public void Path_lookup_finds_a_chrome_binary_on_PATH()
    {
        if (OperatingSystem.IsWindows())
        {
            return; // stub browsers are shell scripts
        }

        var stub = _tmp.WriteStubBrowser("chrome", "exit 0");

        using var reset = new ProbeCacheReset();
        using var noEnv = new EnvVarScope(BrowserLocator.EnvVar, null);
        using var path = new EnvVarScope(
            "PATH", _tmp.Dir + Path.PathSeparator + Environment.GetEnvironmentVariable("PATH"));
        var info = BrowserLocator.Probe(refresh: true);

        Assert.True(info.Found);
        Assert.Equal(stub, info.Path);
        Assert.Equal("chrome", info.Kind);
    }

    [Fact]
    public void Probe_caches_per_process_until_refreshed()
    {
        var first = _tmp.WriteStubBrowser("fake-chrome", "exit 0");
        var second = _tmp.WriteStubBrowser("fake-msedge", "exit 0");

        using var reset = new ProbeCacheReset();
        using (new EnvVarScope(BrowserLocator.EnvVar, first))
        {
            Assert.Equal(first, BrowserLocator.Probe(refresh: true).Path);
        }

        using (new EnvVarScope(BrowserLocator.EnvVar, second))
        {
            // Cached answer survives the env change ...
            Assert.Equal(first, BrowserLocator.Probe().Path);

            // ... until an explicit refresh.
            Assert.Equal(second, BrowserLocator.Probe(refresh: true).Path);
        }
    }
}
