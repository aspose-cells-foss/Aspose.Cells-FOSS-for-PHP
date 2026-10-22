<?php

namespace Aspose\Cells_FOSS\GoldenTests;

require_once __DIR__ . "/../../autoload.php";
require_once __DIR__ . "/../Shared/TestInfrastructure.php";
require_once __DIR__ . "/../Shared/AutoFilterScenarioFactory.php";
require_once __DIR__ . "/../Shared/ChartScenarioFactory.php";
require_once __DIR__ . "/../Shared/ShapeScenarioFactory.php";
require_once __DIR__ . "/../Shared/WorkbookMetadataScenarioFactory.php";

use Aspose\Cells_FOSS\CellValueType;
use Aspose\Cells_FOSS\ChartType;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XmlName;
use Aspose\Cells_FOSS\SaveFormat;
use Aspose\Cells_FOSS\SaveOptions;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\AutoFilterScenarioFactory;
use Aspose\Cells_FOSS\Testing\ChartScenarioFactory;
use Aspose\Cells_FOSS\Testing\ConditionalFormattingScenarioFactory;
use Aspose\Cells_FOSS\Testing\DefinedNameScenarioFactory;
use Aspose\Cells_FOSS\Testing\HyperlinkScenarioFactory;
use Aspose\Cells_FOSS\Testing\PageSetupScenarioFactory;
use Aspose\Cells_FOSS\Testing\StyleScenarioFactory;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\ValidationScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorkbookMetadataScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorkbookScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorksheetScenarioFactory;
use Aspose\Cells_FOSS\Testing\ZipPackageHelper;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    const MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static function Main()
    {
        return TestRunner::Run(
            "CellData.GoldenTests",
            array(
                new TestCase("xlsx_roundtrip_mixed_scalar_cells_file", array(Program::class, "XlsxRoundtripMixedScalarCellsFile")),
                new TestCase("xlsx_roundtrip_mixed_scalar_cells_stream", array(Program::class, "XlsxRoundtripMixedScalarCellsStream")),
                new TestCase("shared_strings_export_uses_sst_when_enabled", array(Program::class, "SharedStringsExportUsesSstWhenEnabled")),
                new TestCase("inline_strings_export_when_shared_strings_disabled", array(Program::class, "InlineStringsExportWhenSharedStringsDisabled")),
                new TestCase("formula_cells_persist_formula_and_cached_value", array(Program::class, "FormulaCellsPersistFormulaAndCachedValue")),
                new TestCase("mac1904_datetime_roundtrip_and_markup", array(Program::class, "Mac1904DateTimeRoundtripAndMarkup")),
                new TestCase("display_runtime_culture_after_roundtrip", array(Program::class, "DisplayRuntimeCultureAfterRoundtrip")),
                new TestCase("styled_cells_roundtrip_and_emit_stylesheet", array(Program::class, "StyledCellsRoundtripAndEmitStylesheet")),
                new TestCase("worksheet_settings_roundtrip_and_emit_expected_markup", array(Program::class, "WorksheetSettingsRoundtripAndEmitExpectedMarkup")),
                new TestCase("worksheet_view_and_tab_color_roundtrip_and_emit_expected_markup", array(Program::class, "WorksheetViewAndTabColorRoundtripAndEmitExpectedMarkup")),
                new TestCase("worksheet_protection_roundtrip_and_emit_expected_markup", array(Program::class, "WorksheetProtectionRoundtripAndEmitExpectedMarkup")),
                new TestCase("autofilter_roundtrip_and_emit_expected_markup", array(Program::class, "AutoFilterRoundtripAndEmitExpectedMarkup")),
                new TestCase("autofilter_omits_invalid_dxf_references", array(Program::class, "AutoFilterOmitsInvalidDxfReferences")),
                new TestCase("defined_names_roundtrip_and_emit_expected_markup", array(Program::class, "DefinedNamesRoundtripAndEmitExpectedMarkup")),
                new TestCase("worksheet_dimension_includes_merge_only_ranges", array(Program::class, "WorksheetDimensionIncludesMergeOnlyRanges")),
                new TestCase("hyperlinks_roundtrip_and_emit_expected_markup", array(Program::class, "HyperlinksRoundtripAndEmitExpectedMarkup")),
                new TestCase("data_validations_roundtrip_and_emit_expected_markup", array(Program::class, "DataValidationsRoundtripAndEmitExpectedMarkup")),
                new TestCase("conditional_formattings_roundtrip_and_emit_expected_markup", array(Program::class, "ConditionalFormattingsRoundtripAndEmitExpectedMarkup")),
                new TestCase("advanced_conditional_formattings_roundtrip_and_emit_expected_markup", array(Program::class, "AdvancedConditionalFormattingsRoundtripAndEmitExpectedMarkup")),
                new TestCase("page_setup_roundtrip_and_emit_expected_markup", array(Program::class, "PageSetupRoundtripAndEmitExpectedMarkup")),
                new TestCase("workbook_metadata_roundtrip_and_emit_expected_markup", array(Program::class, "WorkbookMetadataRoundtripAndEmitExpectedMarkup")),
                new TestCase("extended_document_properties_do_not_inject_default_application_metadata", array(Program::class, "ExtendedDocumentPropertiesDoNotInjectDefaultApplicationMetadata")),
                new TestCase("workbook_metadata_loads_from_root_relationship_targets", array(Program::class, "WorkbookMetadataLoadsFromRootRelationshipTargets")),
                new TestCase("unreferenced_document_properties_parts_are_ignored", array(Program::class, "UnreferencedDocumentPropertiesPartsAreIgnored")),
                new TestCase("shape_file_roundtrip_and_export", array(Program::class, "ShapeFileRoundtripAndExport")),
                new TestCase("chart_file_roundtrip_and_export", array(Program::class, "ChartFileRoundtripAndExport")),
                new TestCase("chart_programmatic_creation_roundtrip", array(Program::class, "ChartProgrammaticCreationRoundtrip")),
                new TestCase("chart_all_types_export", array(Program::class, "ChartAllTypesExport")),
                new TestCase("comment_xml_part_emitted_when_comments_exist", array(Program::class, "CommentXmlPartEmittedWhenCommentsExist")),
                new TestCase("comment_xml_part_not_emitted_when_no_comments", array(Program::class, "CommentXmlPartNotEmittedWhenNoComments")),
                new TestCase("comment_vml_drawing_emitted_when_comments_exist", array(Program::class, "CommentVmlDrawingEmittedWhenCommentsExist")),
                new TestCase("comment_roundtrip_preserves_note_and_author", array(Program::class, "CommentRoundtripPreservesNoteAndAuthor")),
                new TestCase("comment_visible_roundtrip", array(Program::class, "CommentVisibleRoundtrip")),
            )
        );
    }

    public static function XlsxRoundtripMixedScalarCellsFile()
    {
        $temp = new TemporaryDirectory("golden-file");
        $path = $temp->GetPath("mixed.xlsx");
        $workbook = WorkbookScenarioFactory::CreateMixedCellWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        self::AssertMixedWorkbook($loaded, false);
    }

    public static function XlsxRoundtripMixedScalarCellsStream()
    {
        $workbook = WorkbookScenarioFactory::CreateMixedCellWorkbook();
        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $loaded = new Workbook($stream);
        fclose($stream);
        self::AssertMixedWorkbook($loaded, false);
    }

    public static function SharedStringsExportUsesSstWhenEnabled()
    {
        $temp = new TemporaryDirectory("golden-sst");
        $path = $temp->GetPath("shared-strings.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue("Hello");
        $workbook->Worksheets[0]->Cells["A2"]->PutValue("Hello");
        $options = new SaveOptions();
        $options->UseSharedStrings = true;
        $workbook->Save($path, $options);

        AssertEx::True(ZipPackageHelper::EntryExists($path, "xl/sharedStrings.xml"));
        $sharedStringsXml = ZipPackageHelper::ReadEntryText($path, "xl/sharedStrings.xml");
        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");

        AssertEx::Contains("<sst", $sharedStringsXml);
        AssertEx::Contains("uniqueCount=\"1\"", $sharedStringsXml);
        AssertEx::Contains("t=\"s\"", $worksheetXml);
    }

    public static function InlineStringsExportWhenSharedStringsDisabled()
    {
        $temp = new TemporaryDirectory("golden-inline");
        $path = $temp->GetPath("inline.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue("Inline");
        $options = new SaveOptions();
        $options->UseSharedStrings = false;
        $workbook->Save($path, $options);

        AssertEx::False(ZipPackageHelper::EntryExists($path, "xl/sharedStrings.xml"));
        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");

        AssertEx::Contains("t=\"inlineStr\"", $worksheetXml);
        AssertEx::Contains("<t>Inline</t>", $worksheetXml);
    }

    public static function FormulaCellsPersistFormulaAndCachedValue()
    {
        $temp = new TemporaryDirectory("golden-formula");
        $path = $temp->GetPath("formula.xlsx");
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells["A1"]->PutValue(10);
        $sheet->Cells["B1"]->PutValue(20);
        $sheet->Cells["C1"]->PutValue(30);
        $sheet->Cells["C1"]->Formula = "=A1+B1";
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("<f>A1+B1</f>", $worksheetXml);
        AssertEx::Contains("<v>30</v>", $worksheetXml);

        $loaded = new Workbook($path);
        AssertEx::Equal("=A1+B1", $loaded->Worksheets[0]->Cells["C1"]->Formula);
        AssertEx::Equal("30", $loaded->Worksheets[0]->Cells["C1"]->StringValue);
    }

    public static function Mac1904DateTimeRoundtripAndMarkup()
    {
        $temp = new TemporaryDirectory("golden-1904");
        $path = $temp->GetPath("date1904.xlsx");
        $workbook = WorkbookScenarioFactory::CreateMixedCellWorkbook(true);
        $workbook->Save($path);

        $workbookXml = ZipPackageHelper::ReadEntryText($path, "xl/workbook.xml");
        AssertEx::Contains("date1904=\"1\"", $workbookXml);

        $loaded = new Workbook($path);
        AssertEx::True($loaded->Settings->Date1904);
        $expected = self::GetTicks(new \DateTime("2024-05-06 07:08:09", new \DateTimeZone("UTC")));
        $actual = self::GetTicks($loaded->Worksheets[0]->Cells["F1"]->Value);
        AssertEx::Equal($expected, $actual);
    }

    public static function DisplayRuntimeCultureAfterRoundtrip()
    {
        $temp = new TemporaryDirectory("golden-display-culture");
        $path = $temp->GetPath("display-culture.xlsx");
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $timestamp = new \DateTime("2024-05-06 07:08:09");

        $sheet->Cells["A1"]->PutValue(1234.5);
        $numericStyle = $sheet->Cells["A1"]->GetStyle();
        $numericStyle->NumberFormat = "#,##0.00";
        $sheet->Cells["A1"]->SetStyle($numericStyle);

        $sheet->Cells["B1"]->PutValue($timestamp);
        $dateStyle = $sheet->Cells["B1"]->GetStyle();
        $dateStyle->NumberFormat = "[$-409]dddd, mmmm d, yyyy";
        $sheet->Cells["B1"]->SetStyle($dateStyle);

        $workbook->Save($path);

        $loaded = new Workbook($path);
        $frCulture = \Aspose\Cells_FOSS\CultureInfo::GetCultureInfo("fr-FR");
        $enCulture = \Aspose\Cells_FOSS\CultureInfo::GetCultureInfo("en-US");
        $loaded->Settings->Culture = $frCulture;

        $expectedNumeric = "1" . "\u{202F}" . "234,50";
        $expectedDate = "Monday, May 6, 2024";
        AssertEx::Equal($expectedNumeric, $loaded->Worksheets[0]->Cells["A1"]->DisplayStringValue);
        AssertEx::Equal($expectedDate, $loaded->Worksheets[0]->Cells["B1"]->DisplayStringValue);
    }

    public static function StyledCellsRoundtripAndEmitStylesheet()
    {
        $temp = new TemporaryDirectory("golden-styles");
        $path = $temp->GetPath("styles.xlsx");
        $workbook = StyleScenarioFactory::CreateStyledWorkbook();
        $workbook->Save($path);

        AssertEx::True(ZipPackageHelper::EntryExists($path, "xl/styles.xml"));
        $stylesheetXml = ZipPackageHelper::ReadEntryText($path, "xl/styles.xml");
        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");

        AssertEx::Contains("Arial", $stylesheetXml);
        AssertEx::Contains("0.0000", $stylesheetXml);
        AssertEx::Contains("<strike", $stylesheetXml);
        AssertEx::Contains("patternType=\"lightGrid\"", $stylesheetXml);
        AssertEx::Contains("fgColor rgb=\"FFD2DC1E\"", $stylesheetXml);
        AssertEx::Contains("bgColor rgb=\"FF0C2D4E\"", $stylesheetXml);
        AssertEx::Contains("style=\"dotted\"", $stylesheetXml);
        AssertEx::Contains("style=\"mediumDashDot\"", $stylesheetXml);
        AssertEx::Contains("style=\"double\"", $stylesheetXml);
        AssertEx::Contains("style=\"dashDotDot\"", $stylesheetXml);
        AssertEx::Contains("style=\"slantDashDot\"", $stylesheetXml);
        AssertEx::Contains("diagonalUp=\"1\"", $stylesheetXml);
        AssertEx::Contains("diagonalDown=\"1\"", $stylesheetXml);
        AssertEx::Contains("horizontal=\"distributed\"", $stylesheetXml);
        AssertEx::Contains("vertical=\"distributed\"", $stylesheetXml);
        AssertEx::Contains("indent=\"2\"", $stylesheetXml);
        AssertEx::Contains("textRotation=\"45\"", $stylesheetXml);
        AssertEx::Contains("shrinkToFit=\"1\"", $stylesheetXml);
        AssertEx::Contains("readingOrder=\"2\"", $stylesheetXml);
        AssertEx::Contains("relativeIndent=\"1\"", $stylesheetXml);
        AssertEx::Contains("wrapText=\"1\"", $stylesheetXml);
        AssertEx::Contains("locked=\"0\"", $stylesheetXml);
        AssertEx::Contains("hidden=\"1\"", $stylesheetXml);
        AssertEx::Contains("numFmtId=\"4\"", $stylesheetXml);
        self::AssertFontElementOrder($stylesheetXml);
        AssertEx::Contains("s=\"", $worksheetXml);

        $loaded = new Workbook($path);
        StyleScenarioFactory::AssertPrimaryStyle($loaded->Worksheets[0]->Cells["A1"]->GetStyle());
        StyleScenarioFactory::AssertCustomNumberStyle($loaded->Worksheets[0]->Cells["B2"]->GetStyle());
        AssertEx::Equal(CellValueType::IsNull, $loaded->Worksheets[0]->Cells["B2"]->Type);
        AssertEx::Equal(CellValueType::IsDateTime, $loaded->Worksheets[0]->Cells["C3"]->Type);
    }

    private static function AssertFontElementOrder($stylesheetXml)
    {
        $document = XDocument::Load($stylesheetXml);
        $mainNs = "{" . self::MainNs . "}";

        $order = array(
            "b" => 0,
            "i" => 1,
            "strike" => 2,
            "u" => 3,
            "condense" => 4,
            "extend" => 5,
            "outline" => 6,
            "shadow" => 7,
            "charset" => 8,
            "family" => 9,
            "sz" => 10,
            "color" => 11,
            "vertAlign" => 12,
            "scheme" => 13,
            "name" => 14,
        );

        $root = $document->Root;
        $fontsElement = $root === null ? null : $root->Element($mainNs . "fonts");
        $fonts = $fontsElement === null ? array() : $fontsElement->Elements($mainNs . "font");
        foreach ($fonts as $font) {
            $previousOrder = -1;
            foreach ($font->Elements() as $child) {
                $childName = XmlName::split($child->Name)[1];
                AssertEx::True(array_key_exists($childName, $order), "Unexpected font child element: " . $childName);
                $currentOrder = $order[$childName];
                AssertEx::True($currentOrder >= $previousOrder, "Font child element order is invalid: " . $font);
                $previousOrder = $currentOrder;
            }
        }
    }

    public static function WorksheetSettingsRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-worksheet-settings");
        $path = $temp->GetPath("worksheet-settings.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        $workbookXml = ZipPackageHelper::ReadEntryText($path, "xl/workbook.xml");
        AssertEx::Contains("dimension ref=\"A1:C4\"", $worksheetXml);
        AssertEx::Contains("mergeCell ref=\"A1:B2\"", $worksheetXml);
        AssertEx::Contains("hidden=\"1\"", $worksheetXml);
        AssertEx::Contains("state=\"hidden\"", $workbookXml);

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);
        WorksheetScenarioFactory::AssertWorksheetSettingsScenarioHasVisibleSheet($loaded);
    }

    public static function WorksheetViewAndTabColorRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-worksheet-view");
        $path = $temp->GetPath("worksheet-view.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("<sheetViews>", $worksheetXml);
        AssertEx::Contains("showGridLines=\"0\"", $worksheetXml);
        AssertEx::Contains("showRowColHeaders=\"0\"", $worksheetXml);
        AssertEx::Contains("showZeros=\"0\"", $worksheetXml);
        AssertEx::Contains("rightToLeft=\"1\"", $worksheetXml);
        AssertEx::Contains("zoomScale=\"85\"", $worksheetXml);
        AssertEx::Contains("tabColor rgb=\"FF224466\"", $worksheetXml);

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);
        WorksheetScenarioFactory::AssertWorksheetSettingsScenarioHasVisibleSheet($loaded);
    }

    public static function WorksheetProtectionRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-worksheet-protection");
        $path = $temp->GetPath("worksheet-protection.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("<sheetProtection", $worksheetXml);
        AssertEx::Contains("sheet=\"1\"", $worksheetXml);
        AssertEx::Contains("objects=\"1\"", $worksheetXml);
        AssertEx::Contains("scenarios=\"1\"", $worksheetXml);
        AssertEx::Contains("formatCells=\"1\"", $worksheetXml);
        AssertEx::Contains("insertRows=\"1\"", $worksheetXml);
        AssertEx::Contains("autoFilter=\"1\"", $worksheetXml);
        AssertEx::Contains("selectLockedCells=\"1\"", $worksheetXml);
        AssertEx::Contains("selectUnlockedCells=\"1\"", $worksheetXml);

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);
        WorksheetScenarioFactory::AssertWorksheetSettingsScenarioHasVisibleSheet($loaded);
    }

    public static function AutoFilterRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-autofilter");
        $path = $temp->GetPath("autofilter.xlsx");
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        $workbookXml = ZipPackageHelper::ReadEntryText($path, "xl/workbook.xml");
        AssertEx::Contains("<autoFilter ref=\"A1:E6\"", $worksheetXml);
        AssertEx::Contains("<filterColumn colId=\"0\" hiddenButton=\"1\">", $worksheetXml);
        AssertEx::Contains("<filters><filter val=\"Open\" /><filter val=\"Closed\" /></filters>", $worksheetXml);
        AssertEx::Contains("<customFilters and=\"1\">", $worksheetXml);
        AssertEx::Contains("operator=\"greaterThanOrEqual\"", $worksheetXml);
        AssertEx::Contains("operator=\"lessThanOrEqual\"", $worksheetXml);
        AssertEx::Contains("<colorFilter dxfId=\"3\" cellColor=\"1\" />", $worksheetXml);
        AssertEx::Contains("<dynamicFilter type=\"thisMonth\" val=\"1\" maxVal=\"31\" />", $worksheetXml);
        AssertEx::Contains("<top10 top=\"0\" percent=\"1\" val=\"10\" filterVal=\"2.5\" />", $worksheetXml);
        AssertEx::Contains("<sortState ref=\"A2:E6\" caseSensitive=\"1\" sortMethod=\"pinYin\">", $worksheetXml);
        AssertEx::Contains("<sortCondition ref=\"B2:B6\" descending=\"1\" sortBy=\"value\" customList=\"High,Medium,Low\" />", $worksheetXml);
        AssertEx::Contains("<sortCondition ref=\"C2:C6\" sortBy=\"cellColor\" dxfId=\"4\" />", $worksheetXml);
        AssertEx::Contains("<sortCondition ref=\"E2:E6\" sortBy=\"icon\" iconSet=\"3TrafficLights1\" iconId=\"2\" />", $worksheetXml);
        AssertEx::Contains("name=\"_xlnm._FilterDatabase\"", $workbookXml);
        AssertEx::Contains("'Filtered'!\$A\$1:\$E\$6", $workbookXml);

        $loaded = new Workbook($path);
        AutoFilterScenarioFactory::AssertAutoFilter($loaded);
    }

    public static function AutoFilterOmitsInvalidDxfReferences()
    {
        $temp = new TemporaryDirectory("golden-autofilter-invalid-dxf");
        $path = $temp->GetPath("autofilter-invalid-dxf.xlsx");
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = "InvalidDxf";
        $sheet->Cells["A1"]->PutValue("Status");
        $sheet->Cells["A2"]->PutValue("Open");
        $sheet->Cells["A3"]->PutValue("Closed");
        $sheet->Cells["B1"]->PutValue("Amount");
        $sheet->Cells["B2"]->PutValue(10);
        $sheet->Cells["B3"]->PutValue(20);
        $sheet->AutoFilter->Range = "A1:B3";

        $column = $sheet->AutoFilter->FilterColumns[$sheet->AutoFilter->FilterColumns->Add(0)];
        $column->ColorFilter->Enabled = true;
        $column->ColorFilter->DifferentialStyleId = 9;
        $column->ColorFilter->CellColor = true;

        $sheet->AutoFilter->SortState->Ref = "A2:B3";
        $sortCondition = $sheet->AutoFilter->SortState->SortConditions[$sheet->AutoFilter->SortState->SortConditions->Add("B2:B3")];
        $sortCondition->SortBy = "value";
        $sortCondition->DifferentialStyleId = 8;

        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("<autoFilter ref=\"A1:B3\"", $worksheetXml);
        AssertEx::False(strpos($worksheetXml, "<colorFilter") !== false);
        AssertEx::False(strpos($worksheetXml, "dxfId=\"8\"") !== false);
        AssertEx::False(strpos($worksheetXml, "dxfId=\"9\"") !== false);

        $loaded = new Workbook($path);
        AssertEx::Equal("A1:B3", $loaded->Worksheets[0]->AutoFilter->Range);
        AssertEx::Equal(0, $loaded->Worksheets[0]->AutoFilter->FilterColumns->Count);
        AssertEx::Equal(1, $loaded->Worksheets[0]->AutoFilter->SortState->SortConditions->Count);
        AssertEx::Null($loaded->Worksheets[0]->AutoFilter->SortState->SortConditions[0]->DifferentialStyleId);
    }

    public static function DefinedNamesRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-defined-names");
        $path = $temp->GetPath("defined-names.xlsx");
        $workbook = DefinedNameScenarioFactory::CreateDefinedNamesWorkbook();
        $workbook->Save($path);

        $workbookXml = ZipPackageHelper::ReadEntryText($path, "xl/workbook.xml");
        AssertEx::Contains("<definedNames>", $workbookXml);
        AssertEx::Contains("name=\"GlobalRange\"", $workbookXml);
        AssertEx::Contains("hidden=\"1\"", $workbookXml);
        AssertEx::Contains("comment=\"Primary range\"", $workbookXml);
        AssertEx::Contains("name=\"LocalCell\"", $workbookXml);
        AssertEx::Contains("localSheetId=\"1\"", $workbookXml);
        AssertEx::Contains("name=\"_xlnm.Print_Area\"", $workbookXml);
        AssertEx::Contains("name=\"_xlnm.Print_Titles\"", $workbookXml);

        $loaded = new Workbook($path);
        DefinedNameScenarioFactory::AssertDefinedNames($loaded);
    }

    public static function WorksheetDimensionIncludesMergeOnlyRanges()
    {
        $temp = new TemporaryDirectory("golden-merge-dimension");
        $path = $temp->GetPath("merge-only.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Name = "MergeOnly";
        $workbook->Worksheets[0]->Cells->Merge(4, 5, 2, 2);
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("dimension ref=\"F5:G6\"", $worksheetXml);
        AssertEx::Contains("mergeCell ref=\"F5:G6\"", $worksheetXml);

        $loaded = new Workbook($path);
        AssertEx::Equal(1, count($loaded->Worksheets[0]->Cells->MergedCells));
        AssertEx::Equal(4, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartRow);
        AssertEx::Equal(5, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartColumn);
    }

    public static function HyperlinksRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-hyperlinks");
        $path = $temp->GetPath("hyperlinks.xlsx");
        $workbook = HyperlinkScenarioFactory::CreateHyperlinkWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        $relsXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/_rels/sheet1.xml.rels");
        AssertEx::Contains("<hyperlinks>", $worksheetXml);
        AssertEx::Contains("ref=\"A1\"", $worksheetXml);
        AssertEx::Contains("location=\"'Target Sheet'!C3\"", $worksheetXml);
        AssertEx::Contains("Target=\"https://example.com/docs?q=1\"", $relsXml);
        AssertEx::Contains("Target=\"mailto:test@example.com\"", $relsXml);

        $loaded = new Workbook($path);
        HyperlinkScenarioFactory::AssertHyperlinks($loaded);
    }

    public static function DataValidationsRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-validations");
        $path = $temp->GetPath("validations.xlsx");
        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("<dataValidations", $worksheetXml);
        AssertEx::Contains("type=\"list\"", $worksheetXml);
        AssertEx::Contains("type=\"decimal\"", $worksheetXml);
        AssertEx::Contains("type=\"custom\"", $worksheetXml);
        AssertEx::Contains("sqref=\"A1:A3\"", $worksheetXml);
        AssertEx::Contains("sqref=\"B2:C3 E2:E3\"", $worksheetXml);
        AssertEx::Contains("showDropDown=\"1\"", $worksheetXml);
        AssertEx::Contains("<formula1>\"Open,Closed\"</formula1>", $worksheetXml);
        AssertEx::Contains("<formula2>9.5</formula2>", $worksheetXml);

        $loaded = new Workbook($path);
        ValidationScenarioFactory::AssertValidations($loaded);
    }

    public static function ConditionalFormattingsRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-conditional-formatting");
        $path = $temp->GetPath("conditional-formatting.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        $stylesXml = ZipPackageHelper::ReadEntryText($path, "xl/styles.xml");
        AssertEx::Contains("<conditionalFormatting", $worksheetXml);
        AssertEx::Contains("type=\"cellIs\"", $worksheetXml);
        AssertEx::Contains("type=\"expression\"", $worksheetXml);
        AssertEx::Contains("operator=\"between\"", $worksheetXml);
        AssertEx::Contains("stopIfTrue=\"1\"", $worksheetXml);
        AssertEx::Contains("<formula>MOD(A1,2)=0</formula>", $worksheetXml);
        AssertEx::Contains("<dxfs", $stylesXml);
        AssertEx::Contains("count=\"3\"", $stylesXml);

        $loaded = new Workbook($path);
        ConditionalFormattingScenarioFactory::AssertConditionalFormattings($loaded);
    }

    public static function AdvancedConditionalFormattingsRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-advanced-conditional-formatting");
        $path = $temp->GetPath("advanced-conditional-formatting.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateAdvancedConditionalFormattingWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("type=\"containsText\"", $worksheetXml);
        AssertEx::Contains("type=\"notContainsText\"", $worksheetXml);
        AssertEx::Contains("type=\"beginsWith\"", $worksheetXml);
        AssertEx::Contains("type=\"endsWith\"", $worksheetXml);
        AssertEx::Contains("type=\"timePeriod\"", $worksheetXml);
        AssertEx::Contains("type=\"duplicateValues\"", $worksheetXml);
        AssertEx::Contains("type=\"uniqueValues\"", $worksheetXml);
        AssertEx::Contains("type=\"top10\"", $worksheetXml);
        AssertEx::Contains("type=\"aboveAverage\"", $worksheetXml);
        AssertEx::Contains("type=\"colorScale\"", $worksheetXml);
        AssertEx::Contains("type=\"dataBar\"", $worksheetXml);
        AssertEx::Contains("type=\"iconSet\"", $worksheetXml);
        AssertEx::Contains("<colorScale>", $worksheetXml);
        AssertEx::Contains("<dataBar>", $worksheetXml);
        AssertEx::Contains("<iconSet iconSet=\"4Arrows\" reverse=\"1\" showValue=\"0\">", $worksheetXml);

        $loaded = new Workbook($path);
        ConditionalFormattingScenarioFactory::AssertAdvancedConditionalFormattings($loaded);
    }

    public static function PageSetupRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-page-setup");
        $path = $temp->GetPath("page-setup.xlsx");
        $workbook = PageSetupScenarioFactory::CreatePageSetupWorkbook();
        $workbook->Save($path);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/sheet1.xml");
        AssertEx::Contains("pageMargins", $worksheetXml);
        AssertEx::Contains("pageSetup", $worksheetXml);
        AssertEx::Contains("orientation=\"landscape\"", $worksheetXml);
        AssertEx::Contains("paperSize=\"9\"", $worksheetXml);
        AssertEx::Contains("<rowBreaks", $worksheetXml);
        AssertEx::Contains("<colBreaks", $worksheetXml);

        $loaded = new Workbook($path);
        PageSetupScenarioFactory::AssertPageSetup($loaded);
    }

    private static function AssertMixedWorkbook($workbook, $expectedDate1904)
    {
        AssertEx::Equal($expectedDate1904, $workbook->Settings->Date1904);
        AssertEx::Equal("Data", $workbook->Worksheets[0]->Name);
        AssertEx::Equal("Hello", $workbook->Worksheets[0]->Cells["A1"]->StringValue);
        AssertEx::Equal(123, (int) $workbook->Worksheets[0]->Cells["B1"]->Value);
        AssertEx::Equal(true, (bool) $workbook->Worksheets[0]->Cells["C1"]->Value);
        AssertEx::Equal(12.5, (float) $workbook->Worksheets[0]->Cells["D1"]->Value);
        AssertEx::True(abs((float) $workbook->Worksheets[0]->Cells["E1"]->Value - 6.02214076e+23) < 1e+10);
        AssertEx::Equal(
            self::GetTicks(new \DateTime("2024-05-06 07:08:09", new \DateTimeZone("UTC"))),
            self::GetTicks($workbook->Worksheets[0]->Cells["F1"]->Value)
        );
        AssertEx::Equal("=B1*2", $workbook->Worksheets[0]->Cells["G1"]->Formula);
        AssertEx::Equal(20, (int) $workbook->Worksheets[0]->Cells["G1"]->Value);
    }

    public static function WorkbookMetadataRoundtripAndEmitExpectedMarkup()
    {
        $temp = new TemporaryDirectory("golden-workbook-metadata");
        $path = $temp->GetPath("workbook-metadata.xlsx");
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($path);

        $workbookXml = ZipPackageHelper::ReadEntryText($path, "xl/workbook.xml");
        $coreXml = ZipPackageHelper::ReadEntryText($path, "docProps/core.xml");
        $appXml = ZipPackageHelper::ReadEntryText($path, "docProps/app.xml");

        AssertEx::Contains("<workbookPr", $workbookXml);
        AssertEx::Contains("codeName=\"WorkbookCode\"", $workbookXml);
        AssertEx::Contains("showObjects=\"placeholders\"", $workbookXml);
        AssertEx::Contains("<workbookProtection", $workbookXml);
        AssertEx::Contains("workbookPassword=\"ABCD\"", $workbookXml);
        AssertEx::Contains("<bookViews>", $workbookXml);
        AssertEx::Contains("activeTab=\"1\"", $workbookXml);
        AssertEx::Contains("showSheetTabs=\"0\"", $workbookXml);
        AssertEx::Contains("<calcPr", $workbookXml);
        AssertEx::Contains("calcMode=\"manual\"", $workbookXml);
        AssertEx::Contains("refMode=\"R1C1\"", $workbookXml);
        AssertEx::Contains("Quarterly Summary", $coreXml);
        AssertEx::Contains("Automation", $coreXml);
        AssertEx::Contains("Aspose.Cells_FOSS Tests", $appXml);
        AssertEx::Contains("https://example.com/base/", $appXml);

        $loaded = new Workbook($path);
        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($loaded);
    }

    public static function ExtendedDocumentPropertiesDoNotInjectDefaultApplicationMetadata()
    {
        $temp = new TemporaryDirectory("golden-extended-document-properties-no-default-app");
        $path = $temp->GetPath("document-properties-no-default-app.xlsx");
        $workbook = new Workbook();
        $workbook->DocumentProperties->Company = "Aspose Cells FOSS";
        $workbook->DocumentProperties->Manager = "Release";
        $workbook->Save($path);

        $appXml = ZipPackageHelper::ReadEntryText($path, "docProps/app.xml");
        AssertEx::Contains("<Company>Aspose Cells FOSS</Company>", $appXml);
        AssertEx::Contains("<Manager>Release</Manager>", $appXml);
        AssertEx::False(strpos($appXml, "<Application>") !== false);
        AssertEx::False(strpos($appXml, "<AppVersion>") !== false);

        $loaded = new Workbook($path);
        AssertEx::Equal("Aspose Cells FOSS", $loaded->DocumentProperties->Company);
        AssertEx::Equal("Release", $loaded->DocumentProperties->Manager);
        AssertEx::Equal("", $loaded->DocumentProperties->Extended->Application);
        AssertEx::Equal("", $loaded->DocumentProperties->Extended->AppVersion);
    }

    public static function WorkbookMetadataLoadsFromRootRelationshipTargets()
    {
        $temp = new TemporaryDirectory("golden-workbook-metadata-root-relationships");
        $path = $temp->GetPath("workbook-metadata-root-targets.xlsx");
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::MoveEntry($path, "docProps/core.xml", "metadata/core-props.xml");
        ZipPackageHelper::MoveEntry($path, "docProps/app.xml", "metadata/app-props.xml");
        self::RewriteDocumentPropertiesTargets($path, "/metadata/core-props.xml", "/metadata/app-props.xml");

        $loaded = new Workbook($path);
        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($loaded);
    }

    public static function UnreferencedDocumentPropertiesPartsAreIgnored()
    {
        $temp = new TemporaryDirectory("golden-workbook-metadata-orphaned-docprops");
        $path = $temp->GetPath("workbook-metadata-orphaned-docprops.xlsx");
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($path);

        self::RemoveDocumentPropertiesRelationships($path);

        $loaded = new Workbook($path);
        AssertEx::Equal("WorkbookCode", $loaded->Properties->CodeName);
        self::AssertDocumentPropertiesAreDefault($loaded);
    }

    private static function RewriteDocumentPropertiesTargets($packagePath, $coreTarget, $appTarget)
    {
        $contentTypesNs = "{http://schemas.openxmlformats.org/package/2006/content-types}";
        $overrideName = $contentTypesNs . "Override";
        ZipPackageHelper::RewriteXmlEntry($packagePath, "[Content_Types].xml", function ($document) use ($overrideName, $coreTarget, $appTarget) {
            $root = $document->Root;
            $elements = $root === null ? array() : $root->Elements($overrideName);
            foreach ($elements as $element) {
                $partNameAttribute = $element->Attribute("PartName");
                $partName = $partNameAttribute === null ? null : (string) $partNameAttribute;
                if ($partName === "/docProps/core.xml") {
                    $element->SetAttributeValue("PartName", $coreTarget);
                } else if ($partName === "/docProps/app.xml") {
                    $element->SetAttributeValue("PartName", $appTarget);
                }
            }
        });

        $relationshipsNs = "{http://schemas.openxmlformats.org/package/2006/relationships}";
        $relationshipName = $relationshipsNs . "Relationship";
        ZipPackageHelper::RewriteXmlEntry($packagePath, "_rels/.rels", function ($document) use ($relationshipName, $coreTarget, $appTarget) {
            $root = $document->Root;
            $elements = $root === null ? array() : $root->Elements($relationshipName);
            foreach ($elements as $relationship) {
                $typeAttribute = $relationship->Attribute("Type");
                $relationshipType = $typeAttribute === null ? null : (string) $typeAttribute;
                if ($relationshipType === "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties") {
                    $relationship->SetAttributeValue("Target", ltrim($coreTarget, "/"));
                } else if ($relationshipType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties") {
                    $relationship->SetAttributeValue("Target", ltrim($appTarget, "/"));
                }
            }
        });
    }

    private static function RemoveDocumentPropertiesRelationships($packagePath)
    {
        $relationshipsNs = "{http://schemas.openxmlformats.org/package/2006/relationships}";
        $relationshipName = $relationshipsNs . "Relationship";
        ZipPackageHelper::RewriteXmlEntry($packagePath, "_rels/.rels", function ($document) use ($relationshipName) {
            $root = $document->Root;
            $relationships = $root === null ? array() : $root->Elements($relationshipName);
            for ($index = count($relationships) - 1; $index >= 0; $index--) {
                $relationship = $relationships[$index];
                $typeAttribute = $relationship->Attribute("Type");
                $relationshipType = $typeAttribute === null ? null : (string) $typeAttribute;
                if ($relationshipType === "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                    || $relationshipType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties") {
                    $relationship->Remove();
                }
            }
        });
    }

    private static function AssertDocumentPropertiesAreDefault($workbook)
    {
        AssertEx::Equal("", $workbook->DocumentProperties->Title);
        AssertEx::Equal("", $workbook->DocumentProperties->Subject);
        AssertEx::Equal("", $workbook->DocumentProperties->Author);
        AssertEx::Equal("", $workbook->DocumentProperties->Keywords);
        AssertEx::Equal("", $workbook->DocumentProperties->Comments);
        AssertEx::Equal("", $workbook->DocumentProperties->Category);
        AssertEx::Equal("", $workbook->DocumentProperties->Company);
        AssertEx::Equal("", $workbook->DocumentProperties->Manager);
        AssertEx::Equal("", $workbook->DocumentProperties->Core->LastModifiedBy);
        AssertEx::Equal("", $workbook->DocumentProperties->Core->Revision);
        AssertEx::Equal("", $workbook->DocumentProperties->Core->ContentStatus);
        AssertEx::Null($workbook->DocumentProperties->Core->Created);
        AssertEx::Null($workbook->DocumentProperties->Core->Modified);
        AssertEx::Equal("", $workbook->DocumentProperties->Extended->Application);
        AssertEx::Equal("", $workbook->DocumentProperties->Extended->AppVersion);
        AssertEx::Equal(0, $workbook->DocumentProperties->Extended->DocSecurity);
        AssertEx::Equal("", $workbook->DocumentProperties->Extended->HyperlinkBase);
        AssertEx::False((bool) $workbook->DocumentProperties->Extended->ScaleCrop);
        AssertEx::False((bool) $workbook->DocumentProperties->Extended->LinksUpToDate);
        AssertEx::False((bool) $workbook->DocumentProperties->Extended->SharedDoc);
    }

    public static function ShapeFileRoundtripAndExport()
    {
        $temp = new TemporaryDirectory("shape-file-roundtrip");
        $inputPath = self::ResolveRepositoryFile("Input", "shape1.xlsx");
        $outputPath = $temp->GetPath("shape1-export.xlsx");

        $workbook = new Workbook($inputPath);
        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        AssertEx::True(count($loaded->Worksheets[0]->Shapes) > 0, "Exported file should contain shapes");
    }

    public static function ChartFileRoundtripAndExport()
    {
        $inputPath = self::ResolveRepositoryFile("Input", "AreaChart" . DIRECTORY_SEPARATOR . "areachart1.xlsx");
        $outputDir = self::GetOrCreateOutputDirectory("chart-golden");
        $outputPath = $outputDir . DIRECTORY_SEPARATOR . "areachart1-export.xlsx";

        $workbook = new Workbook($inputPath);
        $originalChartCount = count($workbook->Worksheets[0]->Charts);
        AssertEx::True($originalChartCount > 0, "Input file should contain at least one chart");

        $workbook->Save($outputPath);

        $drawingXml = ZipPackageHelper::ReadEntryText($outputPath, "xl/drawings/drawing1.xml");
        AssertEx::Contains("graphicFrame", $drawingXml, "Drawing XML must contain graphicFrame");
        AssertEx::Contains("http://schemas.openxmlformats.org/drawingml/2006/chart", $drawingXml, "Drawing XML must reference chart namespace URI");

        AssertEx::True(ZipPackageHelper::EntryExists($outputPath, "xl/charts/chart1.xml"), "Saved XLSX must contain xl/charts/chart1.xml");

        $contentTypes = ZipPackageHelper::ReadEntryText($outputPath, "[Content_Types].xml");
        AssertEx::Contains("/xl/charts/chart1.xml", $contentTypes, "Content_Types.xml must have chart1.xml override");

        $loaded = new Workbook($outputPath);
        AssertEx::Equal($originalChartCount, count($loaded->Worksheets[0]->Charts), "Chart count must match after round-trip");
    }

    public static function ChartProgrammaticCreationRoundtrip()
    {
        $outputDir = self::GetOrCreateOutputDirectory("chart-golden");
        $outputPath = $outputDir . DIRECTORY_SEPARATOR . "column-chart.xlsx";

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells["A1"]->PutValue("Q1");
        $sheet->Cells["B1"]->PutValue(10);
        $sheet->Cells["A2"]->PutValue("Q2");
        $sheet->Cells["B2"]->PutValue(30);
        $sheet->Cells["A3"]->PutValue("Q3");
        $sheet->Cells["B3"]->PutValue(20);
        $sheet->Charts->Add(ChartType::Column, "Sheet1!\$B\$1:\$B\$3", 5, 0, 20, 10);
        $workbook->Save($outputPath);

        $chartXml = ZipPackageHelper::ReadEntryText($outputPath, "xl/charts/chart1.xml");
        AssertEx::Contains("<c:barChart", $chartXml, "Chart XML must contain barChart element");
        AssertEx::Contains("barDir val=\"col\"", $chartXml, "Chart XML must have col direction for column chart");
        AssertEx::Contains("Sheet1!\$B\$1:\$B\$3", $chartXml, "Chart XML must preserve data range");

        $loaded = new Workbook($outputPath);
        AssertEx::Equal(1, count($loaded->Worksheets[0]->Charts), "Reloaded file must have 1 chart");
        AssertEx::Equal(ChartType::Column, $loaded->Worksheets[0]->Charts[0]->ChartType, "Reloaded chart must be Column type");
    }

    public static function ChartAllTypesExport()
    {
        $outputDir = self::GetOrCreateOutputDirectory("chart-all-types");

        $types = array(
            ChartType::Column, ChartType::Bar, ChartType::Line, ChartType::Area,
            ChartType::Pie, ChartType::Doughnut, ChartType::Scatter, ChartType::Bubble,
            ChartType::Radar, ChartType::Stock,
            ChartType::Column3D, ChartType::Bar3D, ChartType::Line3D, ChartType::Area3D, ChartType::Pie3D,
            ChartType::Surface3D, ChartType::SurfaceWireframe3D, ChartType::Contour,
        );

        foreach ($types as $type) {
            $workbook = ChartScenarioFactory::CreateChartWorkbook($type);
            $outputPath = $outputDir . DIRECTORY_SEPARATOR . self::GetChartTypeName($type) . ".xlsx";
            $workbook->Save($outputPath);

            $reloaded = new Workbook($outputPath);
            AssertEx::Equal(1, count($reloaded->Worksheets[0]->Charts), "Expected 1 chart for type: " . $type);
            AssertEx::Equal($type, $reloaded->Worksheets[0]->Charts[0]->ChartType, "Wrong chart type after export/reload for: " . $type);
        }
    }

    public static function CommentXmlPartEmittedWhenCommentsExist()
    {
        $temp = new TemporaryDirectory("comment-golden");
        $path = $temp->GetPath("comment-exists.xlsx");

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $comment = $sheet->Comments->Add(0, 0);
        $comment->Note = "Hello";
        $comment->Author = "Bob";
        $workbook->Save($path);

        AssertEx::True(ZipPackageHelper::EntryExists($path, "xl/comments1.xml"), "xl/comments1.xml must be present when worksheet has comments");

        $commentsXml = ZipPackageHelper::ReadEntryText($path, "xl/comments1.xml");
        AssertEx::Contains("<comments", $commentsXml, "comments XML must have root element");
        AssertEx::Contains("<author>Bob</author>", $commentsXml, "authors list must contain Bob");
        AssertEx::Contains("Hello", $commentsXml, "comment text must appear in XML");
        AssertEx::Contains("ref=\"A1\"", $commentsXml, "comment ref must be A1");

        $contentTypes = ZipPackageHelper::ReadEntryText($path, "[Content_Types].xml");
        AssertEx::Contains("/xl/comments1.xml", $contentTypes, "[Content_Types].xml must list comments part");
    }

    public static function CommentXmlPartNotEmittedWhenNoComments()
    {
        $temp = new TemporaryDirectory("comment-golden");
        $path = $temp->GetPath("no-comments.xlsx");

        $workbook = new Workbook();
        $workbook->Save($path);

        AssertEx::False(ZipPackageHelper::EntryExists($path, "xl/comments1.xml"), "xl/comments1.xml must NOT be present when worksheet has no comments");
        AssertEx::False(ZipPackageHelper::EntryExists($path, "xl/drawings/vmlDrawing1.vml"), "vmlDrawing1.vml must NOT be present when worksheet has no comments");
    }

    public static function CommentVmlDrawingEmittedWhenCommentsExist()
    {
        $temp = new TemporaryDirectory("comment-golden");
        $path = $temp->GetPath("comment-vml.xlsx");

        $workbook = new Workbook();
        $comment = $workbook->Worksheets[0]->Comments->Add(2, 3);
        $comment->Width = 200;
        $comment->Height = 100;
        $workbook->Save($path);

        AssertEx::True(ZipPackageHelper::EntryExists($path, "xl/drawings/vmlDrawing1.vml"), "vmlDrawing1.vml must be present when worksheet has comments");

        $vmlXml = ZipPackageHelper::ReadEntryText($path, "xl/drawings/vmlDrawing1.vml");
        AssertEx::Contains("v:shape", $vmlXml, "VML must contain v:shape element");
        AssertEx::Contains("ObjectType=\"Note\"", $vmlXml, "VML shape must have ObjectType Note");
        AssertEx::Contains("<x:Row>2</x:Row>", $vmlXml, "VML Row must match comment row");
        AssertEx::Contains("<x:Column>3</x:Column>", $vmlXml, "VML Column must match comment column");

        $contentTypes = ZipPackageHelper::ReadEntryText($path, "[Content_Types].xml");
        AssertEx::Contains("vmlDrawing1.vml", $contentTypes, "[Content_Types].xml must list VML part");

        $sheetRels = ZipPackageHelper::ReadEntryText($path, "xl/worksheets/_rels/sheet1.xml.rels");
        AssertEx::Contains("relationships/comments", $sheetRels, "Sheet rels must contain comments relationship");
        AssertEx::Contains("relationships/vmlDrawing", $sheetRels, "Sheet rels must contain vmlDrawing relationship");
    }

    public static function CommentRoundtripPreservesNoteAndAuthor()
    {
        $temp = new TemporaryDirectory("comment-golden");
        $path = $temp->GetPath("comment-roundtrip.xlsx");

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells["A1"]->PutValue("Data");

        $c1 = $sheet->Comments->Add(0, 0);
        $c1->Note = "First note";
        $c1->Author = "Alice";

        $c2 = $sheet->Comments->Add(1, 2);
        $c2->Note = "Second note";
        $c2->Author = "Bob";

        $workbook->Save($path);

        $loaded = new Workbook($path);
        $loadedSheet = $loaded->Worksheets[0];

        AssertEx::Equal(2, count($loadedSheet->Comments), "Must preserve comment count after roundtrip");

        $found0 = $loadedSheet->Comments["A1"];
        AssertEx::NotNull($found0, "A1 comment must survive roundtrip");
        AssertEx::Equal("First note", $found0->Note);
        AssertEx::Equal("Alice", $found0->Author);

        $found1 = $loadedSheet->Comments["C2"];
        AssertEx::NotNull($found1, "C2 comment must survive roundtrip");
        AssertEx::Equal("Second note", $found1->Note);
        AssertEx::Equal("Bob", $found1->Author);
    }

    public static function CommentVisibleRoundtrip()
    {
        $temp = new TemporaryDirectory("comment-golden");
        $path = $temp->GetPath("comment-visible.xlsx");

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $hidden = $sheet->Comments->Add(0, 0);
        $hidden->Note = "hidden";
        $hidden->IsVisible = false;

        $visible = $sheet->Comments->Add(0, 1);
        $visible->Note = "visible";
        $visible->IsVisible = true;

        $workbook->Save($path);

        $vmlXml = ZipPackageHelper::ReadEntryText($path, "xl/drawings/vmlDrawing1.vml");
        AssertEx::Contains("visibility:visible", $vmlXml, "VML must contain visibility:visible for visible comment");

        $loaded = new Workbook($path);
        $loadedSheet = $loaded->Worksheets[0];

        AssertEx::Equal(2, count($loadedSheet->Comments), "Comment count must be preserved");
        $hiddenLoaded = $loadedSheet->Comments["A1"];
        AssertEx::NotNull($hiddenLoaded);
        AssertEx::False($hiddenLoaded->IsVisible, "Hidden comment must remain hidden after roundtrip");

        $visibleLoaded = $loadedSheet->Comments["B1"];
        AssertEx::NotNull($visibleLoaded);
        AssertEx::True($visibleLoaded->IsVisible, "Visible comment must remain visible after roundtrip");
    }

    private static function GetOrCreateOutputDirectory($subDirectory)
    {
        $directory = __DIR__;
        while (true) {
            if (file_exists($directory . DIRECTORY_SEPARATOR . "composer.json")) {
                $outputDir = $directory . DIRECTORY_SEPARATOR . "output" . DIRECTORY_SEPARATOR . $subDirectory;
                if (!is_dir($outputDir)) {
                    mkdir($outputDir, 0777, true);
                }
                return $outputDir;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }
            $directory = $parent;
        }

        throw new \InvalidArgumentException("Could not find repository root.");
    }

    private static function ResolveRepositoryFile($firstSegment, $secondSegment)
    {
        $directory = __DIR__;
        while (true) {
            if (file_exists($directory . DIRECTORY_SEPARATOR . "composer.json")) {
                return $directory . DIRECTORY_SEPARATOR . $firstSegment . DIRECTORY_SEPARATOR . $secondSegment;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }
            $directory = $parent;
        }

        throw new \InvalidArgumentException("Could not find repository root (composer.json).");
    }

    private static function GetTicks($dateTime)
    {
        if (!($dateTime instanceof \DateTimeInterface)) {
            return null;
        }

        return ($dateTime->getTimestamp() + 62135596800) * 10000000;
    }

    private static function GetChartTypeName($chartType)
    {
        if ($chartType == ChartType::Column) return "Column";
        if ($chartType == ChartType::Bar) return "Bar";
        if ($chartType == ChartType::Line) return "Line";
        if ($chartType == ChartType::Area) return "Area";
        if ($chartType == ChartType::Pie) return "Pie";
        if ($chartType == ChartType::Doughnut) return "Doughnut";
        if ($chartType == ChartType::Scatter) return "Scatter";
        if ($chartType == ChartType::Bubble) return "Bubble";
        if ($chartType == ChartType::Radar) return "Radar";
        if ($chartType == ChartType::Stock) return "Stock";
        if ($chartType == ChartType::Column3D) return "Column3D";
        if ($chartType == ChartType::Bar3D) return "Bar3D";
        if ($chartType == ChartType::Line3D) return "Line3D";
        if ($chartType == ChartType::Area3D) return "Area3D";
        if ($chartType == ChartType::Pie3D) return "Pie3D";
        if ($chartType == ChartType::Surface3D) return "Surface3D";
        if ($chartType == ChartType::SurfaceWireframe3D) return "SurfaceWireframe3D";
        if ($chartType == ChartType::Contour) return "Contour";

        return "Chart" . $chartType;
    }
}

Program::Main();
