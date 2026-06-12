using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

public sealed class SnapshotStoreTests : IDisposable
{
    private readonly string _dir;
    private readonly string _file;
    private readonly SnapshotStore _store;

    public SnapshotStoreTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "aioffice-snap-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
        _file = Path.Combine(_dir, "doc.docx");
        _store = new SnapshotStore(Path.Combine(_dir, "snapshots"));
    }

    public void Dispose()
    {
        if (Directory.Exists(_dir))
        {
            Directory.Delete(_dir, recursive: true);
        }
    }

    [Fact]
    public void Save_then_list_returns_entry_with_matching_rev()
    {
        File.WriteAllText(_file, "v1");
        var saved = _store.Save(_file);

        var listed = _store.List(_file);

        var entry = Assert.Single(listed);
        Assert.Equal(saved.Number, entry.Number);
        Assert.Equal(Rev.OfFile(_file), entry.Rev);
    }

    [Fact]
    public void List_is_empty_for_unknown_file()
    {
        Assert.Empty(_store.List(Path.Combine(_dir, "never-seen.docx")));
    }

    [Fact]
    public void Ring_keeps_only_latest_twenty()
    {
        for (var i = 1; i <= 25; i++)
        {
            File.WriteAllText(_file, $"version {i}");
            _store.Save(_file);
        }

        var listed = _store.List(_file);

        Assert.Equal(SnapshotStore.Capacity, listed.Count);
        Assert.Equal(6, listed[0].Number);   // 1..5 evicted
        Assert.Equal(25, listed[^1].Number); // numbers stay monotonic
    }

    [Fact]
    public void Restore_latest_brings_back_previous_content()
    {
        File.WriteAllText(_file, "original");
        _store.Save(_file);
        File.WriteAllText(_file, "broken edit");

        _store.Restore(_file);

        Assert.Equal("original", File.ReadAllText(_file));
    }

    [Fact]
    public void Restore_specific_number_and_restore_is_undoable()
    {
        File.WriteAllText(_file, "v1");
        var first = _store.Save(_file);
        File.WriteAllText(_file, "v2");
        _store.Save(_file);
        File.WriteAllText(_file, "v3");

        _store.Restore(_file, first.Number);

        Assert.Equal("v1", File.ReadAllText(_file));
        // The pre-restore state ("v3") must itself have been snapshotted.
        Assert.Contains(_store.List(_file), e => e.Rev == Rev.OfString("v3"));
    }

    [Fact]
    public void Restore_unknown_number_lists_candidates()
    {
        File.WriteAllText(_file, "v1");
        _store.Save(_file);

        var ex = Assert.Throws<AiofficeException>(() => _store.Restore(_file, 99));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("1", ex.Candidates);
    }

    [Fact]
    public void Restore_with_no_snapshots_fails_with_suggestion()
    {
        File.WriteAllText(_file, "v1");

        var ex = Assert.Throws<AiofficeException>(() => _store.Restore(_file));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    [Fact]
    public void Save_missing_file_is_file_not_found()
    {
        var ex = Assert.Throws<AiofficeException>(
            () => _store.Save(Path.Combine(_dir, "ghost.docx")));

        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    [Fact]
    public void Rings_are_isolated_per_file_path()
    {
        var other = Path.Combine(_dir, "other.docx");
        File.WriteAllText(_file, "a");
        File.WriteAllText(other, "b");

        _store.Save(_file);

        Assert.Single(_store.List(_file));
        Assert.Empty(_store.List(other));
    }
}
