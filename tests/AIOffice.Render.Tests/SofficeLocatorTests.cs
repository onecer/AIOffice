using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>Restores the soffice probe cache to a clean probe on dispose.</summary>
public sealed class SofficeProbeReset : IDisposable
{
    public void Dispose() => SofficeLocator.Probe(refresh: true);
}

public sealed class SofficeLocatorTests : IDisposable
{
    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    [Fact]
    public void Env_var_with_existing_file_wins_the_probe()
    {
        var stub = _tmp.WriteStubBrowser("fake-soffice", "exit 0");

        using var reset = new SofficeProbeReset();
        using var env = new EnvVarScope(SofficeLocator.EnvVar, stub);
        var info = SofficeLocator.Probe(refresh: true);

        Assert.True(info.Found);
        Assert.Equal(stub, info.Path);
    }

    [Fact]
    public void Env_var_pointing_at_a_missing_file_falls_through()
    {
        var bogus = _tmp.PathOf("nope-soffice");

        using var reset = new SofficeProbeReset();
        using var env = new EnvVarScope(SofficeLocator.EnvVar, bogus);
        var info = SofficeLocator.Probe(refresh: true);

        // The bogus override path must never win, whatever the machine has.
        Assert.NotEqual(bogus, info.Path);
    }

    [Fact]
    public void Path_lookup_finds_a_soffice_binary_on_PATH()
    {
        if (OperatingSystem.IsWindows())
        {
            return; // stub binaries are shell scripts
        }

        var stub = _tmp.WriteStubBrowser("soffice", "exit 0");

        using var reset = new SofficeProbeReset();
        using var noEnv = new EnvVarScope(SofficeLocator.EnvVar, null);
        using var path = new EnvVarScope(
            "PATH", _tmp.Dir + Path.PathSeparator + Environment.GetEnvironmentVariable("PATH"));
        var info = SofficeLocator.Probe(refresh: true);

        Assert.True(info.Found);
        Assert.Equal(stub, info.Path);
    }

    [Fact]
    public void Pdftoppm_path_and_flag_stay_consistent()
    {
        // pdftoppm presence is reported regardless of whether soffice is found
        // (a platform default like the macOS app bundle may still locate soffice).
        using var reset = new SofficeProbeReset();
        var info = SofficeLocator.Probe(refresh: true);

        // The found flag and the path always agree, for both probes.
        Assert.Equal(info.Pdftoppm, info.PdftoppmPath is not null);
        Assert.Equal(info.Found, info.Path is not null);
    }

    [Fact]
    public void NotFound_is_the_empty_value()
    {
        Assert.False(SofficeInfo.NotFound.Found);
        Assert.Null(SofficeInfo.NotFound.Path);
        Assert.False(SofficeInfo.NotFound.Pdftoppm);
        Assert.Null(SofficeInfo.NotFound.PdftoppmPath);
    }
}
