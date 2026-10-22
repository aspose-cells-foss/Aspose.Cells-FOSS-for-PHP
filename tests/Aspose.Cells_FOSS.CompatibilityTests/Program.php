<?php

namespace Aspose\Cells_FOSS\CompatibilityTests;

require_once __DIR__ . "/../../autoload.php";
require_once __DIR__ . "/../Shared/TestInfrastructure.php";
require_once __DIR__ . "/../Shared/AutoFilterScenarioFactory.php";
require_once __DIR__ . "/../Shared/WorkbookMetadataScenarioFactory.php";

use Aspose\Cells_FOSS\AutoShapeType;
use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\CellsException;
use Aspose\Cells_FOSS\CellValueType;
use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\ImageType;
use Aspose\Cells_FOSS\InvalidFileFormatException;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\SaveFormat;
use Aspose\Cells_FOSS\TableStyleType;
use Aspose\Cells_FOSS\TargetModeType;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\VisibilityType;
use Aspose\Cells_FOSS\Workbook;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\AutoFilterScenarioFactory;
use Aspose\Cells_FOSS\Testing\ConditionalFormattingScenarioFactory;
use Aspose\Cells_FOSS\Testing\DefinedNameScenarioFactory;
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

final class Program
{
    public static function Main()
    {
        return TestRunner::Run(
            "CellData.CompatibilityTests",
            array(
                new TestCase("file_and_stream_load_paths_produce_same_values", array(Program::class, "FileAndStreamLoadPathsProduceSameValues")),
                new TestCase("save_overloads_produce_equivalent_workbooks", array(Program::class, "SaveOverloadsProduceEquivalentWorkbooks")),
                new TestCase("formula_setter_accepts_with_or_without_leading_equal", array(Program::class, "FormulaSetterAcceptsWithOrWithoutLeadingEqual")),
                new TestCase("exception_mapping_uses_cells_exception_types", array(Program::class, "ExceptionMappingUsesCellsExceptionTypes")),
                new TestCase("public_type_mapping_matches_after_roundtrip", array(Program::class, "PublicTypeMappingMatchesAfterRoundtrip")),
                new TestCase("value_property_setter_matches_supported_scalar_behavior", array(Program::class, "ValuePropertySetterMatchesSupportedScalarBehavior")),
                new TestCase("compatibility_members_follow_aspose_style", array(Program::class, "CompatibilityMembersFollowAsposeStyle")),
                new TestCase("cellarea_public_members_follow_aspose_style", array(Program::class, "CellAreaPublicMembersFollowAsposeStyle")),
                new TestCase("worksheet_view_members_follow_supported_patterns", array(Program::class, "WorksheetViewMembersFollowSupportedPatterns")),
                new TestCase("worksheet_protection_members_follow_supported_patterns", array(Program::class, "WorksheetProtectionMembersFollowSupportedPatterns")),
                new TestCase("autofilter_members_follow_supported_patterns", array(Program::class, "AutoFilterMembersFollowSupportedPatterns")),
                new TestCase("file_and_stream_roundtrip_preserve_autofilter", array(Program::class, "FileAndStreamRoundtripPreserveAutoFilter")),
                new TestCase("excel_input_autofilter_roundtrip_preserves_header_only_range", array(Program::class, "ExcelInputAutoFilterRoundtripPreservesHeaderOnlyRange")),
                new TestCase("file_and_stream_roundtrip_preserve_defined_names", array(Program::class, "FileAndStreamRoundtripPreserveDefinedNames")),
                new TestCase("file_and_stream_roundtrip_preserve_styles", array(Program::class, "FileAndStreamRoundtripPreserveStyles")),
                new TestCase("file_and_stream_roundtrip_preserve_worksheet_settings", array(Program::class, "FileAndStreamRoundtripPreserveWorksheetSettings")),
                new TestCase("file_and_stream_roundtrip_preserve_data_validations", array(Program::class, "FileAndStreamRoundtripPreserveDataValidations")),
                new TestCase("file_and_stream_roundtrip_preserve_conditional_formattings", array(Program::class, "FileAndStreamRoundtripPreserveConditionalFormattings")),
                new TestCase("file_and_stream_roundtrip_preserve_advanced_conditional_formattings", array(Program::class, "FileAndStreamRoundtripPreserveAdvancedConditionalFormattings")),
                new TestCase("file_and_stream_roundtrip_preserve_page_setup", array(Program::class, "FileAndStreamRoundtripPreservePageSetup")),
                new TestCase("workbook_metadata_members_follow_supported_patterns", array(Program::class, "WorkbookMetadataMembersFollowSupportedPatterns")),
                new TestCase("file_and_stream_roundtrip_preserve_workbook_metadata", array(Program::class, "FileAndStreamRoundtripPreserveWorkbookMetadata")),
                new TestCase("excel_input_table_roundtrip_preserves_table_structure", array(Program::class, "ExcelInputTableRoundtripPreservesTableStructure")),
                new TestCase("excel_input_picture_roundtrip_preserves_picture_structure", array(Program::class, "ExcelInputPictureRoundtripPreservesPictureStructure")),
                new TestCase("excel_input_shape_roundtrip_preserves_shape_structure", array(Program::class, "ExcelInputShapeRoundtripPreservesShapeStructure")),
                new TestCase("excel_input_shape_xlsx_roundtrip_includes_connector", array(Program::class, "ExcelInputShapeXlsxRoundtripIncludesConnector")),
                new TestCase("excel_input_area_chart_roundtrip", array(Program::class, "ExcelInputAreaChartRoundtrip")),
                new TestCase("excel_input_bar_chart_roundtrip", array(Program::class, "ExcelInputBarChartRoundtrip")),
                new TestCase("excel_input_column_chart_roundtrip", array(Program::class, "ExcelInputColumnChartRoundtrip")),
                new TestCase("excel_input_line_chart_roundtrip", array(Program::class, "ExcelInputLineChartRoundtrip")),
                new TestCase("excel_input_pie_chart_roundtrip", array(Program::class, "ExcelInputPieChartRoundtrip")),
                new TestCase("excel_input_xy_chart_roundtrip", array(Program::class, "ExcelInputXyChartRoundtrip")),
                new TestCase("excel_input_radar_chart_roundtrip", array(Program::class, "ExcelInputRadarChartRoundtrip")),
                new TestCase("excel_input_surface_chart_roundtrip", array(Program::class, "ExcelInputSurfaceChartRoundtrip")),
                new TestCase("excel_input_stock_chart_roundtrip", array(Program::class, "ExcelInputStockChartRoundtrip")),
                new TestCase("excel_input_funnel_chart_roundtrip", array(Program::class, "ExcelInputFunnelChartRoundtrip")),
                new TestCase("excel_input_histogram_chart_roundtrip", array(Program::class, "ExcelInputHistogramChartRoundtrip")),
                new TestCase("excel_input_box_chart_roundtrip", array(Program::class, "ExcelInputBoxChartRoundtrip")),
                new TestCase("excel_input_sunburst_chart_roundtrip", array(Program::class, "ExcelInputSunburstChartRoundtrip")),
                new TestCase("excel_input_treemap_chart_roundtrip", array(Program::class, "ExcelInputTreemapChartRoundtrip")),
                new TestCase("excel_input_map_chart_roundtrip", array(Program::class, "ExcelInputMapChartRoundtrip")),
                new TestCase("excel_input_combo_chart_roundtrip", array(Program::class, "ExcelInputComboChartRoundtrip")),
                new TestCase("excel_input_waterfall_chart_roundtrip", array(Program::class, "ExcelInputWaterfallChartRoundtrip")),
                new TestCase("excel_input_allcharts_roundtrip", array(Program::class, "ExcelInputAllChartsRoundtrip")),
                new TestCase("excel_input_sparkline_roundtrip_no_crash", array(Program::class, "ExcelInputSparklineRoundtripNoCrash")),
                new TestCase("comment_programmatic_roundtrip_preserves_count", array(Program::class, "CommentProgrammaticRoundtripPreservesCount")),
                new TestCase("comment_programmatic_roundtrip_preserves_note_and_author", array(Program::class, "CommentProgrammaticRoundtripPreservesNoteAndAuthor"))
            ));
    }

    public static function FileAndStreamLoadPathsProduceSameValues()
    {
        $temp = new TemporaryDirectory("compat-load-paths");
        $path = $temp->GetPath("book.xlsx");
        $workbook = WorkbookScenarioFactory::CreateMixedCellWorkbook();
        $workbook->Save($path);

        $stream = fopen($path, "rb");
        $fromFile = new Workbook($path);
        $fromStream = new Workbook($stream);
        fclose($stream);

        Program::AssertWorkbookDataEquals($fromFile, $fromStream);
    }

    public static function SaveOverloadsProduceEquivalentWorkbooks()
    {
        $temp = new TemporaryDirectory("compat-save-overloads");
        $filePath = $temp->GetPath("book-file.xlsx");
        $workbook = WorkbookScenarioFactory::CreateMixedCellWorkbook();
        $workbook->Save($filePath, SaveFormat::Xlsx);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        Program::AssertWorkbookDataEquals($fromFile, $fromStream);
    }

    public static function FormulaSetterAcceptsWithOrWithoutLeadingEqual()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells["A1"];
        $cell->PutValue(10);
        $cell->Formula = "B1+C1";
        AssertEx::Equal("=B1+C1", $cell->Formula);

        $cell->Formula = "=D1+E1";
        AssertEx::Equal("=D1+E1", $cell->Formula);
    }

    public static function ExceptionMappingUsesCellsExceptionTypes()
    {
        AssertEx::Throws(CellsException::class, function () {
            $unused = (new Workbook())->Worksheets["missing"];
        });
        AssertEx::Throws(CellsException::class, function () {
            $unused = (new Workbook())->Worksheets[0]->Cells["1A"];
        });
        AssertEx::Throws(InvalidFileFormatException::class, function () {
            $stream = fopen("php://temp", "r+b");
            fwrite($stream, "\x01\x02\x03\x04");
            rewind($stream);
            $unused = new Workbook($stream);
            fclose($stream);
        });
    }

    public static function PublicTypeMappingMatchesAfterRoundtrip()
    {
        $temp = new TemporaryDirectory("compat-types");
        $path = $temp->GetPath("types.xlsx");
        $workbook = WorkbookScenarioFactory::CreateMixedCellWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        AssertEx::Equal(CellValueType::IsString, $loaded->Worksheets[0]->Cells["A1"]->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $loaded->Worksheets[0]->Cells["B1"]->Type);
        AssertEx::Equal(CellValueType::IsBool, $loaded->Worksheets[0]->Cells["C1"]->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $loaded->Worksheets[0]->Cells["D1"]->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $loaded->Worksheets[0]->Cells["E1"]->Type);
        AssertEx::Equal(CellValueType::IsDateTime, $loaded->Worksheets[0]->Cells["F1"]->Type);
        AssertEx::Equal(CellValueType::IsUnknown, $loaded->Worksheets[0]->Cells["G1"]->Type);
    }

    public static function ValuePropertySetterMatchesSupportedScalarBehavior()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells["A1"]->Value = "alpha";
        $sheet->Cells["B1"]->Value = 12;
        $sheet->Cells["C1"]->Value = true;
        $sheet->Cells["D1"]->Value = new \DateTime("2024-01-02 03:04:00");
        $sheet->Cells["E1"]->Value = null;

        AssertEx::Equal("alpha", $sheet->Cells["A1"]->Value);
        AssertEx::Equal(12, (int)$sheet->Cells["B1"]->Value);
        AssertEx::Equal(true, (bool)$sheet->Cells["C1"]->Value);
        AssertEx::Equal(CellValueType::IsDateTime, $sheet->Cells["D1"]->Type);
        AssertEx::Equal("", $sheet->Cells["E1"]->DisplayStringValue);
    }

    public static function CompatibilityMembersFollowAsposeStyle()
    {
        $workbook = new Workbook();
        $workbook->Settings->Date1904 = true;
        AssertEx::True($workbook->Settings->Date1904);

        $index = $workbook->Worksheets->Add();
        $workbook->Worksheets[$index]->Name = "Report";
        $workbook->Worksheets->ActiveSheetName = "Report";
        AssertEx::Equal($index, $workbook->Worksheets->ActiveSheetIndex);

        $sheet = $workbook->Worksheets[$index];
        $sheet->VisibilityType = VisibilityType::Hidden;
        AssertEx::Equal(VisibilityType::Hidden, $sheet->VisibilityType);

        $sheet->Cells["A1"]->Value = "Docs";
        $hyperlinkIndex = $sheet->Hyperlinks->Add(0, 0, 1, 1, "https://example.com/docs");
        $hyperlink = $sheet->Hyperlinks[$hyperlinkIndex];
        AssertEx::Equal(TargetModeType::External, $hyperlink->LinkType);
        $hyperlink->Delete();
        AssertEx::Equal(0, $sheet->Hyperlinks->Count);

        $validationIndex = $sheet->Validations->Add(CellArea::CreateCellArea("B2", "B2"));
        $validation = $sheet->Validations[$validationIndex];
        $validation->Type = ValidationType::List;
        $validation->Formula1 = "\"Yes,No\"";
        AssertEx::Equal(ValidationType::List, $sheet->Validations->GetValidationInCell(1, 1)->Type);

        $conditionalFormattingIndex = $sheet->ConditionalFormattings->Add();
        $conditionalFormatting = $sheet->ConditionalFormattings[$conditionalFormattingIndex];
        $conditionalFormatting->AddArea(CellArea::CreateCellArea("C3", "C3"));
        $conditionalFormatting->AddCondition(FormatConditionType::Expression, OperatorType::None, "C3>0", "");
        AssertEx::Equal(1, $sheet->ConditionalFormattings->Count);

        $pageSetup = $sheet->PageSetup;
        $pageSetup->LeftMargin = 1.27;
        AssertEx::Equal(0.5, $pageSetup->LeftMarginInch);
    }

    public static function CellAreaPublicMembersFollowAsposeStyle()
    {
        $cellAreaType = new \ReflectionClass(CellArea::class);

        AssertEx::True($cellAreaType->hasProperty("StartRow"));
        AssertEx::True($cellAreaType->hasProperty("EndRow"));
        AssertEx::True($cellAreaType->hasProperty("StartColumn"));
        AssertEx::True($cellAreaType->hasProperty("EndColumn"));

        AssertEx::False($cellAreaType->hasProperty("FirstRow"));
        AssertEx::False($cellAreaType->hasProperty("FirstColumn"));
        AssertEx::False($cellAreaType->hasProperty("TotalRows"));
        AssertEx::False($cellAreaType->hasProperty("TotalColumns"));

        // CellArea must not be constructible with start/end row and column indices; use CreateCellArea.
        $constructor = $cellAreaType->getConstructor();
        $constructorTakesEndIndices = false;
        if ($constructor !== null) {
            foreach ($constructor->getParameters() as $parameter) {
                if ($parameter->getName() === "endRow") {
                    $constructorTakesEndIndices = true;
                }
            }
        }
        AssertEx::False($constructorTakesEndIndices);

        AssertEx::True($cellAreaType->hasMethod("CreateCellArea"));
        $createCellArea = $cellAreaType->getMethod("CreateCellArea");
        AssertEx::True($createCellArea->isStatic());
        AssertEx::Equal(4, $createCellArea->getNumberOfParameters());
    }

    public static function WorksheetViewMembersFollowSupportedPatterns()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->TabColor = Color::FromArgb(255, 34, 68, 102);
        $sheet->ShowGridlines = false;
        $sheet->ShowRowColumnHeaders = false;
        $sheet->ShowZeros = false;
        $sheet->RightToLeft = true;
        $sheet->Zoom = 85;

        AssertEx::Equal(Color::FromArgb(255, 34, 68, 102), $sheet->TabColor);
        AssertEx::False($sheet->ShowGridlines);
        AssertEx::False($sheet->ShowRowColumnHeaders);
        AssertEx::False($sheet->ShowZeros);
        AssertEx::True($sheet->RightToLeft);
        AssertEx::Equal(85, $sheet->Zoom);
    }

    public static function WorksheetProtectionMembersFollowSupportedPatterns()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->Protect();
        $sheet->Protection->Objects = true;
        $sheet->Protection->Scenarios = true;
        $sheet->Protection->AutoFilter = true;
        $sheet->Protection->SelectLockedCells = true;
        $sheet->Protection->SelectUnlockedCells = true;

        AssertEx::True($sheet->Protection->IsProtected);
        AssertEx::True($sheet->Protection->Objects);
        AssertEx::True($sheet->Protection->Scenarios);
        AssertEx::True($sheet->Protection->AutoFilter);
        AssertEx::True($sheet->Protection->SelectLockedCells);
        AssertEx::True($sheet->Protection->SelectUnlockedCells);
    }

    public static function AutoFilterMembersFollowSupportedPatterns()
    {
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        AutoFilterScenarioFactory::AssertAutoFilter($workbook);
    }

    public static function FileAndStreamRoundtripPreserveAutoFilter()
    {
        $temp = new TemporaryDirectory("compat-autofilter");
        $filePath = $temp->GetPath("autofilter.xlsx");
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        AutoFilterScenarioFactory::AssertAutoFilter($fromFile);
        AutoFilterScenarioFactory::AssertAutoFilter($fromStream);
    }

    public static function FileAndStreamRoundtripPreserveDefinedNames()
    {
        $temp = new TemporaryDirectory("compat-defined-names");
        $filePath = $temp->GetPath("defined-names.xlsx");
        $workbook = DefinedNameScenarioFactory::CreateDefinedNamesWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        DefinedNameScenarioFactory::AssertDefinedNames($fromFile);
        DefinedNameScenarioFactory::AssertDefinedNames($fromStream);
    }

    public static function FileAndStreamRoundtripPreserveStyles()
    {
        $temp = new TemporaryDirectory("compat-styles");
        $filePath = $temp->GetPath("styled-file.xlsx");
        $workbook = StyleScenarioFactory::CreateStyledWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        StyleScenarioFactory::AssertPrimaryStyle($fromFile->Worksheets[0]->Cells["A1"]->GetStyle());
        StyleScenarioFactory::AssertPrimaryStyle($fromStream->Worksheets[0]->Cells["A1"]->GetStyle());
        StyleScenarioFactory::AssertCustomNumberStyle($fromFile->Worksheets[0]->Cells["B2"]->GetStyle());
        StyleScenarioFactory::AssertCustomNumberStyle($fromStream->Worksheets[0]->Cells["B2"]->GetStyle());
        AssertEx::Equal(CellValueType::IsNull, $fromFile->Worksheets[0]->Cells["B2"]->Type);
        AssertEx::Equal(CellValueType::IsNull, $fromStream->Worksheets[0]->Cells["B2"]->Type);
    }

    public static function FileAndStreamRoundtripPreserveWorksheetSettings()
    {
        $temp = new TemporaryDirectory("compat-worksheet-settings");
        $filePath = $temp->GetPath("worksheet-settings.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        WorksheetScenarioFactory::AssertWorksheetSettings($fromFile);
        WorksheetScenarioFactory::AssertWorksheetSettingsScenarioHasVisibleSheet($fromFile);
        WorksheetScenarioFactory::AssertWorksheetSettings($fromStream);
        WorksheetScenarioFactory::AssertWorksheetSettingsScenarioHasVisibleSheet($fromStream);
    }

    public static function FileAndStreamRoundtripPreserveDataValidations()
    {
        $temp = new TemporaryDirectory("compat-validations");
        $filePath = $temp->GetPath("validations.xlsx");
        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        ValidationScenarioFactory::AssertValidations($fromFile);
        ValidationScenarioFactory::AssertValidations($fromStream);
    }

    public static function FileAndStreamRoundtripPreserveConditionalFormattings()
    {
        $temp = new TemporaryDirectory("compat-conditional-formatting");
        $filePath = $temp->GetPath("conditional-formatting.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        ConditionalFormattingScenarioFactory::AssertConditionalFormattings($fromFile);
        ConditionalFormattingScenarioFactory::AssertConditionalFormattings($fromStream);
    }

    public static function FileAndStreamRoundtripPreserveAdvancedConditionalFormattings()
    {
        $temp = new TemporaryDirectory("compat-advanced-conditional-formatting");
        $filePath = $temp->GetPath("advanced-conditional-formatting.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateAdvancedConditionalFormattingWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        ConditionalFormattingScenarioFactory::AssertAdvancedConditionalFormattings($fromFile);
        ConditionalFormattingScenarioFactory::AssertAdvancedConditionalFormattings($fromStream);
    }

    private static function AssertWorkbookDataEquals(Workbook $expected, Workbook $actual)
    {
        $expectedSheet = $expected->Worksheets[0];
        $actualSheet = $actual->Worksheets[0];

        foreach (array("A1", "B1", "C1", "D1", "E1", "F1", "G1") as $cellName) {
            AssertEx::Equal($expectedSheet->Cells[$cellName]->Type, $actualSheet->Cells[$cellName]->Type, "Type mismatch for " . $cellName . ".");
            AssertEx::Equal($expectedSheet->Cells[$cellName]->StringValue, $actualSheet->Cells[$cellName]->StringValue, "Value mismatch for " . $cellName . ".");
            AssertEx::Equal($expectedSheet->Cells[$cellName]->Formula, $actualSheet->Cells[$cellName]->Formula, "Formula mismatch for " . $cellName . ".");
        }
    }

    public static function FileAndStreamRoundtripPreservePageSetup()
    {
        $temp = new TemporaryDirectory("compat-page-setup");
        $filePath = $temp->GetPath("page-setup.xlsx");
        $workbook = PageSetupScenarioFactory::CreatePageSetupWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        PageSetupScenarioFactory::AssertPageSetup($fromFile);
        PageSetupScenarioFactory::AssertPageSetup($fromStream);
    }

    public static function WorkbookMetadataMembersFollowSupportedPatterns()
    {
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($workbook);
        AssertEx::Equal("Data", $workbook->Worksheets[$workbook->Properties->View->ActiveTab]->Name);
    }

    public static function FileAndStreamRoundtripPreserveWorkbookMetadata()
    {
        $temp = new TemporaryDirectory("compat-workbook-metadata");
        $filePath = $temp->GetPath("workbook-metadata.xlsx");
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($filePath);

        $stream = fopen("php://temp", "r+b");
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);

        $fromFile = new Workbook($filePath);
        $fromStream = new Workbook($stream);
        fclose($stream);

        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($fromFile);
        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($fromStream);
    }

    public static function ExcelInputAutoFilterRoundtripPreservesHeaderOnlyRange()
    {
        $inputPath = Program::ResolveRepositoryFile("Input", "Autofilter.xlsx");
        $workbook = new Workbook($inputPath);
        AssertEx::Equal("A2:C2", $workbook->Worksheets[0]->AutoFilter->Range);
        AssertEx::Equal(0, $workbook->Worksheets[0]->AutoFilter->FilterColumns->Count);
        AssertEx::Equal(0, $workbook->Worksheets[0]->AutoFilter->SortState->SortConditions->Count);

        $temp = new TemporaryDirectory("compat-input-autofilter");
        $roundtripPath = $temp->GetPath("autofilter-roundtrip.xlsx");
        $workbook->Save($roundtripPath);

        $reloaded = new Workbook($roundtripPath);
        AssertEx::Equal("A2:C2", $reloaded->Worksheets[0]->AutoFilter->Range);
        AssertEx::Equal(0, $reloaded->Worksheets[0]->AutoFilter->FilterColumns->Count);
        AssertEx::Equal(0, $reloaded->Worksheets[0]->AutoFilter->SortState->SortConditions->Count);

        $workbookXml = ZipPackageHelper::ReadEntryText($roundtripPath, "xl/workbook.xml");
        $worksheetXml = ZipPackageHelper::ReadEntryText($roundtripPath, "xl/worksheets/sheet1.xml");
        AssertEx::Contains('<autoFilter ref="A2:C2"', $worksheetXml);
        AssertEx::Contains('name="_xlnm._FilterDatabase"', $workbookXml);
        AssertEx::Contains('$A$2:$C$2', $workbookXml);
    }

    public static function ExcelInputTableRoundtripPreservesTableStructure()
    {
        $inputPath = Program::ResolveRepositoryFile("Input", "table.xlsx");
        $workbook = new Workbook($inputPath);
        $sheet = $workbook->Worksheets[0];

        AssertEx::Equal(1, $sheet->ListObjects->Count);
        $table = $sheet->ListObjects[0];
        AssertEx::Equal("Table5", $table->DisplayName);
        AssertEx::Equal(0, $table->StartRow);
        AssertEx::Equal(0, $table->StartColumn);
        AssertEx::Equal(3, $table->EndRow);
        AssertEx::Equal(2, $table->EndColumn);
        AssertEx::True($table->ShowHeaderRow);
        AssertEx::False($table->ShowTotals);
        AssertEx::Equal(TableStyleType::TableStyleMedium2, $table->TableStyleType);
        AssertEx::True($table->ShowTableStyleRowStripes);
        AssertEx::False($table->ShowTableStyleFirstColumn);
        AssertEx::False($table->ShowTableStyleLastColumn);
        AssertEx::False($table->ShowTableStyleColumnStripes);
        AssertEx::Equal(3, $table->ListColumns->Count);
        AssertEx::Equal("Product", $table->ListColumns[0]->Name);
        AssertEx::Equal("Region", $table->ListColumns[1]->Name);
        AssertEx::Equal("Revenue", $table->ListColumns[2]->Name);

        $temp = new TemporaryDirectory("compat-input-table");
        $roundtripPath = $temp->GetPath("table-roundtrip.xlsx");
        $workbook->Save($roundtripPath);

        $reloaded = new Workbook($roundtripPath);
        $reloadedTable = $reloaded->Worksheets[0]->ListObjects[0];
        AssertEx::Equal("Table5", $reloadedTable->DisplayName);
        AssertEx::Equal(0, $reloadedTable->StartRow);
        AssertEx::Equal(3, $reloadedTable->EndRow);
        AssertEx::False($reloadedTable->ShowTotals);
        AssertEx::Equal(TableStyleType::TableStyleMedium2, $reloadedTable->TableStyleType);
        AssertEx::Equal(3, $reloadedTable->ListColumns->Count);
        AssertEx::Equal("Product", $reloadedTable->ListColumns[0]->Name);
        AssertEx::Equal("Revenue", $reloadedTable->ListColumns[2]->Name);

        $tableXml = ZipPackageHelper::ReadEntryText($roundtripPath, "xl/tables/table1.xml");
        AssertEx::Contains('ref="A1:C4"', $tableXml);
        AssertEx::Contains('totalsRowShown="0"', $tableXml);
        AssertEx::Contains('name="TableStyleMedium2"', $tableXml);
    }

    public static function ExcelInputPictureRoundtripPreservesPictureStructure()
    {
        $inputPath = Program::ResolveRepositoryFile("Input", "compare" . DIRECTORY_SEPARATOR . "17.6.3-ByCells.xlsx");
        $workbook = new Workbook($inputPath);
        $sheet = $workbook->Worksheets[0];

        AssertEx::True($sheet->Pictures->Count > 0, "Expected at least one picture in the input workbook.");
        $picture = $sheet->Pictures[0];
        AssertEx::True($picture->Data !== null && strlen($picture->Data) > 0, "Picture data must be non-empty.");
        AssertEx::Equal(ImageType::Jpeg, $picture->ImageType);
        AssertEx::True($picture->UpperLeftRow >= 0, "UpperLeftRow must be non-negative.");
        AssertEx::True($picture->LowerRightRow >= $picture->UpperLeftRow, "LowerRightRow must be >= UpperLeftRow.");

        $originalDataLength = strlen($picture->Data);
        $originalUpperLeftRow = $picture->UpperLeftRow;
        $originalUpperLeftColumn = $picture->UpperLeftColumn;
        $originalLowerRightRow = $picture->LowerRightRow;
        $originalLowerRightColumn = $picture->LowerRightColumn;

        $temp = new TemporaryDirectory("compat-input-picture");
        $roundtripPath = $temp->GetPath("picture-roundtrip.xlsx");
        $workbook->Save($roundtripPath);

        $reloaded = new Workbook($roundtripPath);
        $reloadedSheet = $reloaded->Worksheets[0];
        AssertEx::Equal($sheet->Pictures->Count, $reloadedSheet->Pictures->Count);

        $reloadedPicture = $reloadedSheet->Pictures[0];
        AssertEx::Equal($originalDataLength, strlen($reloadedPicture->Data));
        AssertEx::Equal($originalUpperLeftRow, $reloadedPicture->UpperLeftRow);
        AssertEx::Equal($originalUpperLeftColumn, $reloadedPicture->UpperLeftColumn);
        AssertEx::Equal($originalLowerRightRow, $reloadedPicture->LowerRightRow);
        AssertEx::Equal($originalLowerRightColumn, $reloadedPicture->LowerRightColumn);
        AssertEx::Equal(ImageType::Jpeg, $reloadedPicture->ImageType);

        $drawingXml = ZipPackageHelper::ReadEntryText($roundtripPath, "xl/drawings/drawing1.xml");
        AssertEx::Contains("<xdr:twoCellAnchor", $drawingXml);
        AssertEx::Contains("<xdr:pic>", $drawingXml);
        AssertEx::Contains("rId1", $drawingXml);
    }

    public static function ExcelInputShapeRoundtripPreservesShapeStructure()
    {
        $inputPath = Program::ResolveRepositoryFile("Input", "compare" . DIRECTORY_SEPARATOR . "100_rows.xlsx");
        $workbook = new Workbook($inputPath);
        $sheet = $workbook->Worksheets[0];

        AssertEx::Equal(6, $sheet->Shapes->Count);

        $rect = $sheet->Shapes[0];
        AssertEx::Equal("Rectangle 1", $rect->Name);
        AssertEx::Equal(AutoShapeType::Rectangle, $rect->AutoShapeType);
        AssertEx::True($rect->UpperLeftRow >= 0, "UpperLeftRow must be non-negative.");
        AssertEx::True($rect->LowerRightRow >= $rect->UpperLeftRow, "LowerRightRow must be >= UpperLeftRow.");

        $arrow = $sheet->Shapes[2];
        AssertEx::Equal(AutoShapeType::RightArrow, $arrow->AutoShapeType);

        $temp = new TemporaryDirectory("compat-input-shape");
        $roundtripPath = $temp->GetPath("shape-roundtrip.xlsx");
        $workbook->Save($roundtripPath);

        $reloaded = new Workbook($roundtripPath);
        $reloadedSheet = $reloaded->Worksheets[0];
        AssertEx::Equal(6, $reloadedSheet->Shapes->Count);
        AssertEx::Equal("Rectangle 1", $reloadedSheet->Shapes[0]->Name);
        AssertEx::Equal(AutoShapeType::Rectangle, $reloadedSheet->Shapes[0]->AutoShapeType);
        AssertEx::Equal(AutoShapeType::RightArrow, $reloadedSheet->Shapes[2]->AutoShapeType);

        $drawingXml = ZipPackageHelper::ReadEntryText($roundtripPath, "xl/drawings/drawing1.xml");
        AssertEx::Contains("<xdr:sp ", $drawingXml);
        AssertEx::Contains("Rectangle 1", $drawingXml);
        AssertEx::Contains('prst="rect"', $drawingXml);
        AssertEx::Contains('prst="rightArrow"', $drawingXml);
    }

    public static function ExcelInputShapeXlsxRoundtripIncludesConnector()
    {
        $inputPath = Program::ResolveRepositoryFile("Input", "shape.xlsx");
        $workbook = new Workbook($inputPath);
        $sheet = $workbook->Worksheets[0];

        // shape.xlsx has 3 drawing objects: 1 cxnSp connector + 2 sp shapes
        AssertEx::Equal(3, $sheet->Shapes->Count);

        // Connector is first in document order
        AssertEx::Equal("Straight Arrow Connector 2", $sheet->Shapes[0]->Name);
        AssertEx::Equal(AutoShapeType::StraightConnector, $sheet->Shapes[0]->AutoShapeType);

        // Regular sp shapes
        AssertEx::Equal("Arrow: Down 3", $sheet->Shapes[1]->Name);
        AssertEx::Equal(AutoShapeType::DownArrow, $sheet->Shapes[1]->AutoShapeType);

        AssertEx::Equal("Plus Sign 4", $sheet->Shapes[2]->Name);
        AssertEx::Equal(AutoShapeType::MathPlus, $sheet->Shapes[2]->AutoShapeType);

        $temp = new TemporaryDirectory("compat-input-shape-xlsx");
        $roundtripPath = $temp->GetPath("shape-roundtrip.xlsx");
        $workbook->Save($roundtripPath);

        $reloaded = new Workbook($roundtripPath);
        $reloadedSheet = $reloaded->Worksheets[0];
        AssertEx::Equal(3, $reloadedSheet->Shapes->Count);
        AssertEx::Equal("Straight Arrow Connector 2", $reloadedSheet->Shapes[0]->Name);
        AssertEx::Equal(AutoShapeType::StraightConnector, $reloadedSheet->Shapes[0]->AutoShapeType);
        AssertEx::Equal("Arrow: Down 3", $reloadedSheet->Shapes[1]->Name);
        AssertEx::Equal(AutoShapeType::DownArrow, $reloadedSheet->Shapes[1]->AutoShapeType);
        AssertEx::Equal("Plus Sign 4", $reloadedSheet->Shapes[2]->Name);
        AssertEx::Equal(AutoShapeType::MathPlus, $reloadedSheet->Shapes[2]->AutoShapeType);

        $drawingXml = ZipPackageHelper::ReadEntryText($roundtripPath, "xl/drawings/drawing1.xml");
        AssertEx::Contains("<xdr:cxnSp", $drawingXml);
        AssertEx::Contains("<xdr:sp ", $drawingXml);
        AssertEx::Contains("downArrow", $drawingXml);
        AssertEx::Contains("mathPlus", $drawingXml);
    }

    public static function ExcelInputAreaChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("AreaChart", true);
    }

    public static function ExcelInputBarChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("BarChart", true);
    }

    public static function ExcelInputColumnChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("ColumnChart", true);
    }

    public static function ExcelInputLineChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("linechart", true);
    }

    public static function ExcelInputPieChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("PieChart", true);
    }

    public static function ExcelInputXyChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("XYChart", true);
    }

    public static function ExcelInputRadarChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("RadarChart", true);
    }

    public static function ExcelInputSurfaceChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("SurfaceChart", true);
    }

    public static function ExcelInputStockChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("StockChart", true);
    }

    public static function ExcelInputFunnelChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("FunnelChart", false);
    }

    public static function ExcelInputHistogramChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("HistogramChart", false);
    }

    public static function ExcelInputBoxChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("BoxChart", false);
    }

    public static function ExcelInputSunburstChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("SunburstChart", false);
    }

    public static function ExcelInputTreemapChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("TreemapChart", false);
    }

    public static function ExcelInputMapChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("MapChart", false);
    }

    public static function ExcelInputComboChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("combochart", false);
    }

    public static function ExcelInputWaterfallChartRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("waterfallchart", false);
    }

    public static function ExcelInputAllChartsRoundtrip()
    {
        Program::RoundtripAllChartFilesInDirectory("allcharts", false);
    }

    public static function ExcelInputSparklineRoundtripNoCrash()
    {
        Program::RoundtripAllChartFilesInDirectory("Sparkline", false);
    }

    public static function CommentProgrammaticRoundtripPreservesCount()
    {
        $outputDir = Program::GetOrCreateOutputDirectory("comment-roundtrips");
        $outputPath = $outputDir . DIRECTORY_SEPARATOR . "comment-count.xlsx";

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Comments->Add(0, 0)->Note = "First";
        $sheet->Comments->Add(1, 1)->Note = "Second";
        $sheet->Comments->Add(2, 0)->Note = "Third";
        $workbook->Save($outputPath);

        $reloaded = new Workbook($outputPath);
        AssertEx::Equal(3, $reloaded->Worksheets[0]->Comments->Count,
            "Comment count must be preserved after save/load roundtrip");
    }

    public static function CommentProgrammaticRoundtripPreservesNoteAndAuthor()
    {
        $outputDir = Program::GetOrCreateOutputDirectory("comment-roundtrips");
        $outputPath = $outputDir . DIRECTORY_SEPARATOR . "comment-content.xlsx";

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $c = $sheet->Comments->Add(3, 5);
        $c->Note = "Important: review before merge";
        $c->Author = "Carol";
        $c->IsVisible = true;
        $c->Width = 180;
        $c->Height = 90;
        $workbook->Save($outputPath);

        $reloaded = new Workbook($outputPath);
        $loadedSheet = $reloaded->Worksheets[0];

        AssertEx::Equal(1, $loadedSheet->Comments->Count);
        $loaded = $loadedSheet->Comments["F4"];
        AssertEx::NotNull($loaded, "Comment at F4 must survive roundtrip");
        AssertEx::Equal("Important: review before merge", $loaded->Note);
        AssertEx::Equal("Carol", $loaded->Author);
        AssertEx::True($loaded->IsVisible, "IsVisible must be preserved");
        AssertEx::Equal(180, $loaded->Width, "Width must be preserved");
        AssertEx::Equal(90, $loaded->Height, "Height must be preserved");
    }

    private static function RoundtripAllChartFilesInDirectory($subDirectory, $expectAtLeastOneChart)
    {
        $inputDir = Program::ResolveRepositoryFile("Input", $subDirectory);
        if (!is_dir($inputDir)) {
            throw new \RuntimeException("Input directory not found: " . $inputDir);
        }

        $files = glob($inputDir . DIRECTORY_SEPARATOR . "*.xlsx");
        if ($files === false || count($files) === 0) {
            throw new \RuntimeException("No .xlsx files found in: " . $inputDir);
        }

        $outputDir = Program::GetOrCreateOutputDirectory("chart-roundtrips" . DIRECTORY_SEPARATOR . $subDirectory);

        foreach ($files as $filePath) {
            $fileName = basename($filePath);
            $original = new Workbook($filePath);
            $sheet = $original->Worksheets[0];
            $originalChartCount = $sheet->Charts->Count;

            if ($expectAtLeastOneChart) {
                AssertEx::True($originalChartCount > 0, "Expected at least one chart in: " . $fileName);
            }

            $outputPath = $outputDir . DIRECTORY_SEPARATOR . $fileName;
            $original->Save($outputPath);

            $reloaded = new Workbook($outputPath);
            AssertEx::Equal($originalChartCount, $reloaded->Worksheets[0]->Charts->Count,
                "Chart count changed after round-trip for: " . $fileName);
        }
    }

    private static function GetOrCreateOutputDirectory($subDirectory)
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== "") {
            if (is_file($directory . DIRECTORY_SEPARATOR . "composer.json")) {
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

        throw new \RuntimeException("Could not locate repository root.");
    }

    private static function ResolveRepositoryFile($firstSegment, $secondSegment)
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== "") {
            if (is_file($directory . DIRECTORY_SEPARATOR . "composer.json")) {
                return $directory . DIRECTORY_SEPARATOR . $firstSegment . DIRECTORY_SEPARATOR . $secondSegment;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }

            $directory = $parent;
        }

        throw new \RuntimeException("Could not locate repository root.");
    }
}

Program::Main();
