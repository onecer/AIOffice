using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.2 form-control slice: checkbox | optionButton | spinner | comboBox |
/// listBox | button, authored as legacy form controls on raw OpenXml (VML
/// drawing + ctrlProp + worksheet controls) in a post-save pass. Raw assertions
/// reopen the file with the OpenXml SDK so the control writer cannot grade its
/// own homework, and a ClosedXML round-trip proves the parts survive its saves.
/// </summary>
public sealed class FormControlTests : ExcelTestBase
{
    /// <summary>The worksheet's controls (AlternateContent) entries plus its ctrlProp parts.</summary>
    private static (S.Worksheet Worksheet, WorksheetPart Part) RawSheet(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return (worksheetPart.Worksheet!, worksheetPart);
    }

    private static int ControlCount(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (worksheet, _) = RawSheet(document);
        var controls = worksheet.Elements<S.Controls>().FirstOrDefault();
        return controls?.Elements<DocumentFormat.OpenXml.AlternateContent>().Count() ?? 0;
    }

    [Fact]
    public void Add_checkbox_writes_control_vml_and_ctrlprop_validator_clean()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var (worksheet, part) = RawSheet(document);
            // The worksheet references a legacy drawing and carries one control.
            Assert.True(worksheet.Elements<S.LegacyDrawing>().Any());
            Assert.Single(worksheet.Elements<S.Controls>().First()
                .Elements<DocumentFormat.OpenXml.AlternateContent>());
            // A VML drawing and a ctrlProp part both exist.
            Assert.True(part.VmlDrawingParts.Any());
            var ctrlProps = part.GetPartsOfType<ControlPropertiesPart>().ToList();
            Assert.Single(ctrlProps);
            Assert.Equal(X14.ObjectTypeValues.CheckBox, ctrlProps[0].FormControlProperties!.ObjectType!.Value);
        }

        AssertValidatorClean(file);
    }

    [Theory]
    [InlineData("checkbox", "checkbox")]
    [InlineData("optionButton", "optionButton")]
    [InlineData("button", "button")]
    public void Simple_control_kinds_roundtrip_with_linked_cell(string kind, string expectedKind)
    {
        var file = CreateWorkbook();

        var props = kind == "button"
            ? new (string, JsonNode?)[] { ("kind", kind), ("cell", "E2"), ("text", "Run") }
            : new (string, JsonNode?)[] { ("kind", kind), ("cell", "E2"), ("linkedCell", "F2") };
        var envelope = EditOps(file, AddOp("/Sheet1", "formControl", props));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/formControl[1]"))));
        Assert.Equal(expectedKind, data["controlKind"]!.GetValue<string>());
        Assert.Equal("E2", data["cell"]!.GetValue<string>());
        if (kind != "button")
        {
            Assert.Equal("$F$2", data["linkedCell"]!.GetValue<string>());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Spinner_stores_min_max_increment_and_linked_cell()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "formControl",
                ("kind", "spinner"), ("cell", "E4"), ("linkedCell", "F4"),
                ("min", 0), ("max", 10), ("increment", 2)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var (_, part) = RawSheet(document);
            var fcp = part.GetPartsOfType<ControlPropertiesPart>().First().FormControlProperties!;
            Assert.Equal(X14.ObjectTypeValues.Spin, fcp.ObjectType!.Value);
            Assert.Equal(0u, fcp.Min!.Value);
            Assert.Equal(10u, fcp.Max!.Value);
            Assert.Equal(2u, fcp.Incremental!.Value);
            Assert.Equal("$F$4", fcp.FmlaLink!.Value);
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/formControl[1]"))));
        Assert.Equal(0, data["min"]!.GetValue<int>());
        Assert.Equal(10, data["max"]!.GetValue<int>());
        Assert.Equal(2, data["increment"]!.GetValue<int>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Combo_box_with_list_fill_range_roundtrips()
    {
        var file = CreateWorkbook();
        // The source list lives in H1:H3.
        EditOps(file, SetOp("/Sheet1/H1:H3", ("values", new JsonArray(
            new JsonArray("Red"), new JsonArray("Green"), new JsonArray("Blue")))));

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "formControl",
                ("kind", "comboBox"), ("cell", "E6"), ("linkedCell", "F6"), ("listFillRange", "H1:H3")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var (_, part) = RawSheet(document);
            var fcp = part.GetPartsOfType<ControlPropertiesPart>().First().FormControlProperties!;
            Assert.Equal(X14.ObjectTypeValues.Drop, fcp.ObjectType!.Value);
            Assert.Equal("$H$1:$H$3", fcp.FmlaRange!.Value);
            Assert.Equal("$F$6", fcp.FmlaLink!.Value);
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/formControl[1]"))));
        Assert.Equal("comboBox", data["controlKind"]!.GetValue<string>());
        Assert.Equal("E6", data["cell"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void List_box_with_items_is_accepted()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "formControl",
                ("kind", "listBox"), ("cell", "E8"), ("linkedCell", "F8"),
                ("items", new JsonArray("A", "B", "C"))));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/formControl[1]"))));
        Assert.Equal("listBox", data["controlKind"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Multiple_controls_index_per_sheet_and_list_in_structure()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")),
            AddOp("/Sheet1", "formControl", ("kind", "optionButton"), ("cell", "E3"), ("linkedCell", "F3")),
            AddOp("/Sheet1", "formControl", ("kind", "spinner"), ("cell", "E4"), ("linkedCell", "F4"),
                ("min", 0), ("max", 5)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var controls = structure["sheets"]![0]!["formControls"]!.AsArray();
        Assert.Equal(3, controls.Count);
        Assert.Equal("/Sheet1/formControl[1]", controls[0]!["path"]!.GetValue<string>());
        Assert.Equal("checkbox", controls[0]!["kind"]!.GetValue<string>());
        Assert.Equal("E2", controls[0]!["cell"]!.GetValue<string>());
        Assert.Equal("spinner", controls[2]!["kind"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Remove_form_control_drops_it_and_renumbers()
    {
        var file = CreateWorkbook();
        EditOps(
            file,
            AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")),
            AddOp("/Sheet1", "formControl", ("kind", "optionButton"), ("cell", "E3"), ("linkedCell", "F3")));
        Assert.Equal(2, ControlCount(file));

        var envelope = EditOps(file, RemoveOp("/Sheet1/formControl[1]"));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal(1, ControlCount(file));
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/formControl[1]"))));
        Assert.Equal("optionButton", data["controlKind"]!.GetValue<string>()); // the survivor took index 1
        AssertValidatorClean(file);
    }

    [Fact]
    public void Form_controls_survive_a_later_unrelated_edit()
    {
        var file = CreateWorkbook();
        EditOps(file, AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")));

        // A later value edit forces a ClosedXML re-save; the control (and its
        // anchor cell) must survive intact.
        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", "hello")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal(1, ControlCount(file));
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/formControl[1]"))));
        Assert.Equal("E2", data["cell"]!.GetValue<string>()); // anchor read from the durable VML <x:Anchor>
        Assert.Equal("$F$2", data["linkedCell"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Form_control_opens_through_closedxml_round_trip()
    {
        var file = CreateWorkbook();
        EditOps(file, AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")));

        // ClosedXML must be able to open the file we wrote (it does not author
        // controls, but must preserve them) — proves no corruption.
        using (var workbook = new XLWorkbook(file))
        {
            Assert.Equal("Sheet1", workbook.Worksheets.First().Name);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Unknown_kind_is_unsupported_feature_with_candidates()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "formControl", ("kind", "toggle"), ("cell", "E2")));
        Assert.False(envelope.IsOk);
        var error = Json(envelope)["error"]!;
        Assert.Equal(ErrorCodes.UnsupportedFeature, error["code"]!.GetValue<string>());
        Assert.NotNull(error["candidates"]);
        Assert.False(string.IsNullOrEmpty(error["suggestion"]!.GetValue<string>()));
    }

    [Fact]
    public void Combo_without_items_or_range_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "formControl", ("kind", "comboBox"), ("cell", "E2")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Spinner_with_inverted_range_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "formControl",
                ("kind", "spinner"), ("cell", "E2"), ("min", 10), ("max", 5)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Form_control_on_a_cell_target_path_is_invalid()
    {
        var file = CreateWorkbook();

        // The op must target a sheet path; a cell path is rejected.
        var envelope = EditOps(file, AddOp("/Sheet1/E2", "formControl", ("kind", "checkbox"), ("cell", "E2")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_nonexistent_form_control_is_invalid_path_with_candidates()
    {
        var file = CreateWorkbook();
        EditOps(file, AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")));

        var envelope = EditOps(file, RemoveOp("/Sheet1/formControl[5]"));
        Assert.False(envelope.IsOk);
        var error = Json(envelope)["error"]!;
        Assert.Equal(ErrorCodes.InvalidPath, error["code"]!.GetValue<string>());
        Assert.NotNull(error["candidates"]);
    }

    [Fact]
    public void Set_on_a_form_control_is_unsupported_feature()
    {
        var file = CreateWorkbook();
        EditOps(file, AddOp("/Sheet1", "formControl", ("kind", "checkbox"), ("cell", "E2"), ("linkedCell", "F2")));

        var envelope = EditOps(file, SetOp("/Sheet1/formControl[1]", ("value", 1)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }
}
