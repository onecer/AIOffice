using AIOffice.Core;
using AIOffice.Core.Cli;
using Xunit;

namespace AIOffice.Core.Tests;

public class ArgParserTests
{
    [Fact]
    public void Parses_verb_positionals_and_options()
    {
        var args = ArgParser.Parse(["edit", "report.docx", "--set", "/body/p[1]", "--dry-run", "--json"]);

        Assert.Equal("edit", args.Verb);
        Assert.Equal(["report.docx"], args.Positionals);
        Assert.Equal("/body/p[1]", args.GetOption("set"));
        Assert.True(args.HasFlag("dry-run"));
        Assert.True(args.HasFlag("json"));
    }

    [Fact]
    public void Supports_equals_syntax_and_short_flags()
    {
        var args = ArgParser.Parse(["render", "deck.pptx", "--to=svg", "-o", "out.svg"]);

        Assert.Equal("svg", args.GetOption("to"));
        Assert.Equal("out.svg", args.GetOption("o"));
    }

    [Fact]
    public void Repeated_options_collect_all_values_and_last_wins()
    {
        var args = ArgParser.Parse(["query", "--where", "a", "--where", "b"]);

        Assert.Equal(["a", "b"], args.GetOptionValues("where"));
        Assert.Equal("b", args.GetOption("where"));
    }

    [Fact]
    public void Double_dash_ends_option_parsing()
    {
        var args = ArgParser.Parse(["read", "--", "--weird-file-name.docx"]);

        Assert.Equal("read", args.Verb);
        Assert.Equal(["--weird-file-name.docx"], args.Positionals);
    }

    [Fact]
    public void Known_boolean_flags_do_not_swallow_the_next_token()
    {
        var args = ArgParser.Parse(["read", "--json", "report.docx"]);

        Assert.Equal("true", args.GetOption("json"));
        Assert.Equal(["report.docx"], args.Positionals);
    }

    [Fact]
    public void Negative_numbers_are_values_not_flags()
    {
        var args = ArgParser.Parse(["edit", "--indent", "-2"]);

        Assert.Equal("-2", args.GetOption("indent"));
    }

    [Fact]
    public void Missing_option_returns_null_and_empty_values()
    {
        var args = ArgParser.Parse(["doctor"]);

        Assert.Null(args.GetOption("nope"));
        Assert.Empty(args.GetOptionValues("nope"));
        Assert.False(args.HasFlag("nope"));
    }

    [Fact]
    public void ExpandAtFile_reads_file_through_sandbox()
    {
        var dir = Path.Combine(Path.GetTempPath(), "aioffice-arg-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            File.WriteAllText(Path.Combine(dir, "ops.json"), "[{\"op\":\"set\"}]");
            var workspace = new Workspace(dir);

            Assert.Equal("[{\"op\":\"set\"}]", ArgParser.ExpandAtFile("@ops.json", workspace));
            Assert.Equal("plain", ArgParser.ExpandAtFile("plain", workspace));
            Assert.Equal("@literal", ArgParser.ExpandAtFile("@@literal", workspace));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void ExpandAtFile_denies_paths_outside_sandbox()
    {
        var dir = Path.Combine(Path.GetTempPath(), "aioffice-arg-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var workspace = new Workspace(dir);
            var ex = Assert.Throws<AiofficeException>(
                () => ArgParser.ExpandAtFile("@../../etc/hosts", workspace));

            Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void EditOp_batch_parses_and_validates()
    {
        var ops = EditOp.ParseBatch(
            "[{\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"Hello\"}}," +
            "{\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"p\",\"position\":\"after\"}]");

        Assert.Equal(2, ops.Count);
        Assert.Equal("set", ops[0].Op);
        Assert.Equal("Hello", ops[0].Props!["text"]!.GetValue<string>());
        Assert.Equal("after", ops[1].Position);
    }

    [Theory]
    [InlineData("not json")]
    [InlineData("[]")]
    [InlineData("[{\"op\":\"explode\",\"path\":\"/body/p[1]\"}]")]
    [InlineData("[{\"op\":\"set\"}]")]
    [InlineData("[{\"op\":\"set\",\"path\":\"no-slash\"}]")]
    public void EditOp_batch_rejects_invalid_input(string json)
    {
        var ex = Assert.Throws<AiofficeException>(() => EditOp.ParseBatch(json));

        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    [Fact]
    public void HandlerRegistry_resolves_by_extension_and_rejects_unknown()
    {
        var registry = new HandlerRegistry();
        registry.Register(new FakeHandler(), "docx");

        Assert.Equal(DocumentKind.Docx, registry.Resolve("/tmp/letter.DOCX").Kind);

        var ex = Assert.Throws<AiofficeException>(() => registry.Resolve("/tmp/notes.txt"));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.NotNull(ex.Candidates);
    }

    private sealed class FakeHandler : IFormatHandler
    {
        public DocumentKind Kind => DocumentKind.Docx;

        private static Envelope NotImplemented() => Envelope.Fail(
            ErrorCodes.UnsupportedFeature, "fake", "use a real handler");

        public Envelope Create(CommandContext ctx) => NotImplemented();
        public Envelope Read(CommandContext ctx) => NotImplemented();
        public Envelope Get(CommandContext ctx) => NotImplemented();
        public Envelope Query(CommandContext ctx) => NotImplemented();
        public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops) => NotImplemented();
        public Envelope Render(CommandContext ctx) => NotImplemented();
        public Envelope Validate(CommandContext ctx) => NotImplemented();
        public Envelope Template(CommandContext ctx) => NotImplemented();
    }
}
