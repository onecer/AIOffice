using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

public sealed class RenderValidateTemplateTests : ExcelTestBase
{
    [Fact]
    public void Render_html_applies_number_formats()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "Total"), ("bold", true)),
            SetOp("/Sheet1/B1", ("value", 1234.5), ("numberFormat", "#,##0.0"))).IsOk);

        var data = OkData(Handler.Render(Ctx(file, ("to", "html"))));

        var content = data["content"]!.GetValue<string>();
        Assert.Contains("<table", content, StringComparison.Ordinal);
        Assert.Contains("<strong>Total</strong>", content, StringComparison.Ordinal);
        Assert.Contains("1,234.5", content, StringComparison.Ordinal); // format applied, not raw 1234.5
    }

    [Fact]
    public void Render_html_cells_carry_data_aio_path()
    {
        // The M1 render contract: every cell maps a browser click back to a
        // canonical document path via data-aio-path.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "x")),
            SetOp("/Sheet1/B2", ("value", 5))).IsOk);

        var data = OkData(Handler.Render(Ctx(file, ("to", "html"))));

        var content = data["content"]!.GetValue<string>();
        Assert.Contains("<table data-sheet=\"Sheet1\" data-aio-path=\"/Sheet1\">", content, StringComparison.Ordinal);
        Assert.Contains("<td data-aio-path=\"/Sheet1/A1\">", content, StringComparison.Ordinal);
        Assert.Contains("<td data-aio-path=\"/Sheet1/B2\">", content, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_png_is_unsupported_with_a_named_workaround()
    {
        var file = CreateWorkbook();

        var envelope = Handler.Render(Ctx(file, ("to", "png")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("html", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Validate_reports_formula_errors_as_warnings_not_failures()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "=1/0"))).IsOk);

        var data = OkData(Handler.Validate(Ctx(file)));

        Assert.True(data["valid"]!.GetValue<bool>()); // schema-valid
        Assert.Equal(0, data["errors"]!.GetValue<int>());
        var issue = Assert.Single(data["issues"]!.AsArray());
        Assert.Equal("formula_error", issue!["code"]!.GetValue<string>());
        Assert.Contains("#DIV/0!", issue["message"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Validate_clean_workbook_reports_no_issues()
    {
        var file = CreateWorkbook();

        var data = OkData(Handler.Validate(Ctx(file)));

        Assert.True(data["valid"]!.GetValue<bool>());
        Assert.Empty(data["issues"]!.AsArray());
    }

    [Fact]
    public void Template_fills_placeholders_typed_and_inline()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "Hello {{name}}, total {{total}}")),
            SetOp("/Sheet1/B1", ("value", "{{total}}"), ("valueType", "text"))).IsOk);

        var dataArg = new JsonObject { ["name"] = "Ann", ["total"] = 42 };
        var envelope = Handler.Template(Ctx(file, ("data", dataArg)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var a1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("Hello Ann, total 42", a1["value"]!.GetValue<string>());

        // A whole-cell placeholder takes the JSON value TYPED.
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal("number", b1["type"]!.GetValue<string>());
        Assert.Equal(42.0, b1["value"]!.GetValue<double>());

        // Two snapshots: one from the setup edit, one from in-place templating —
        // template is undoable just like edit.
        Assert.Equal(2, Snapshots.List(file).Count);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Template_warns_about_unresolved_placeholders()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "Hi {{name}} from {{city}}"))).IsOk);

        var envelope = Handler.Template(Ctx(file, ("data", new JsonObject { ["name"] = "Ann" })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var warning = Assert.Single(envelope.Meta.Warnings!);
        Assert.Equal("template_unresolved", warning.Code);
        Assert.Contains("city", warning.Message, StringComparison.Ordinal);

        var a1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("Hi Ann from {{city}}", a1["value"]!.GetValue<string>());
    }

    [Fact]
    public void Template_can_write_to_a_separate_output_file()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "{{x}}"))).IsOk);
        var revBefore = Rev.OfFile(file);

        var envelope = Handler.Template(Ctx(
            file,
            ("data", new JsonObject { ["x"] = 1 }),
            ("output", "filled.xlsx")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(revBefore, Rev.OfFile(file)); // source untouched
        var output = Path.Combine(Dir, "filled.xlsx");
        Assert.True(File.Exists(output));
        AssertValidatorClean(output);
    }
}
