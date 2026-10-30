<?php

namespace Aspose\Cells_FOSS\UnitTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';
require_once __DIR__ . '/../Shared/AutoFilterScenarioFactory.php';
require_once __DIR__ . '/../Shared/ChartScenarioFactory.php';
require_once __DIR__ . '/../Shared/PictureScenarioFactory.php';
require_once __DIR__ . '/../Shared/WorkbookMetadataScenarioFactory.php';

use Aspose\Cells_FOSS\AutoShapeType;
use Aspose\Cells_FOSS\BorderStyleType;
use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\CellValueType;
use Aspose\Cells_FOSS\CellsException;
use Aspose\Cells_FOSS\ChartType;
use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\Font;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\HorizontalAlignmentType;
use Aspose\Cells_FOSS\ImageType;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\PageOrientationType;
use Aspose\Cells_FOSS\PaperSizeType;
use Aspose\Cells_FOSS\Style;
use Aspose\Cells_FOSS\StyleFlag;
use Aspose\Cells_FOSS\TableStyleType;
use Aspose\Cells_FOSS\TotalsCalculation;
use Aspose\Cells_FOSS\UnsupportedFeatureException;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\VerticalAlignmentType;
use Aspose\Cells_FOSS\VisibilityType;
use Aspose\Cells_FOSS\Workbook;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\AutoFilterScenarioFactory;
use Aspose\Cells_FOSS\Testing\ChartScenarioFactory;
use Aspose\Cells_FOSS\Testing\PictureScenarioFactory;
use Aspose\Cells_FOSS\Testing\StyleScenarioFactory;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\WorkbookMetadataScenarioFactory;

final class Program
{
    public static function Main()
    {
        return TestRunner::Run(
            'CellData.UnitTests',
            array(
                new TestCase('a1_indexers_roundtrip', array(Program::class, 'A1IndexersRoundtrip')),
                new TestCase('cells_style_property_roundtrip', array(Program::class, 'CellsStylePropertyRoundtrip')),
                new TestCase('put_value_overloads_assign_expected_types', array(Program::class, 'PutValueOverloadsAssignExpectedTypes')),
                new TestCase('put_value_string_and_object_overloads_follow_aspose_shape', array(Program::class, 'PutValueStringAndObjectOverloadsFollowAsposeShape')),
                new TestCase('stringvalue_formats_supported_scalar_types', array(Program::class, 'StringValueFormatsSupportedScalarTypes')),
                new TestCase('displaystringvalue_applies_numeric_formats_and_stringvalue_stays_raw', array(Program::class, 'DisplayStringValueAppliesNumericFormatsAndStringValueStaysRaw')),
                new TestCase('displaystringvalue_applies_date_and_text_formats', array(Program::class, 'DisplayStringValueAppliesDateAndTextFormats')),
                new TestCase('displaystringvalue_uses_workbook_culture_and_locale_directives', array(Program::class, 'DisplayStringValueUsesWorkbookCultureAndLocaleDirectives')),
                new TestCase('displaystringvalue_uses_english_month_abbreviations_without_locale_directive', array(Program::class, 'DisplayStringValueUsesEnglishMonthAbbreviationsWithoutLocaleDirective')),
                new TestCase('displaystringvalue_applies_extended_date_tokens', array(Program::class, 'DisplayStringValueAppliesExtendedDateTokens')),
                new TestCase('displaystringvalue_applies_long_time_and_elapsed_fraction_with_culture', array(Program::class, 'DisplayStringValueAppliesLongTimeAndElapsedFractionWithCulture')),
                new TestCase('displaystringvalue_elapsed_time_honors_1904_date_system', array(Program::class, 'DisplayStringValueElapsedTimeHonors1904DateSystem')),
                new TestCase('formula_property_normalizes_and_preserves_cached_value', array(Program::class, 'FormulaPropertyNormalizesAndPreservesCachedValue')),
                new TestCase('blank_cells_are_blank_by_default', array(Program::class, 'BlankCellsAreBlankByDefault')),
                new TestCase('worksheet_name_and_collection_guards', array(Program::class, 'WorksheetNameAndCollectionGuards')),
                new TestCase('style_mutation_requires_setstyle_and_returns_clones', array(Program::class, 'StyleMutationRequiresSetStyleAndReturnsClones')),
                new TestCase('style_api_covers_all_public_settings', array(Program::class, 'StyleApiCoversAllPublicSettings')),
                new TestCase('font_equals_and_tostring_follow_aspose_shape', array(Program::class, 'FontEqualsAndToStringFollowAsposeShape')),
                new TestCase('style_equals_hashcode_are_consistent', array(Program::class, 'StyleEqualsHashCodeAreConsistent')),
                new TestCase('style_copy_copies_all_public_settings', array(Program::class, 'StyleCopyCopiesAllPublicSettings')),
                new TestCase('cell_getstyle_bool_overload_matches_getstyle', array(Program::class, 'CellGetStyleBoolOverloadMatchesGetStyle')),
                new TestCase('cell_setstyle_bool_overload_matches_setstyle', array(Program::class, 'CellSetStyleBoolOverloadMatchesSetStyle')),
                new TestCase('cell_setstyle_styleflag_applies_selected_properties', array(Program::class, 'CellSetStyleStyleFlagAppliesSelectedProperties')),
                new TestCase('cell_setstyle_styleflag_all_applies_full_style', array(Program::class, 'CellSetStyleStyleFlagAllAppliesFullStyle')),
                new TestCase('cell_setstyle_styleflag_alignments_applies_indent_settings', array(Program::class, 'CellSetStyleStyleFlagAlignmentsAppliesIndentSettings')),
                new TestCase('worksheet_row_column_and_merge_apis_mutate_expected_settings', array(Program::class, 'WorksheetRowColumnAndMergeApisMutateExpectedSettings')),
                new TestCase('worksheet_view_apis_mutate_expected_settings', array(Program::class, 'WorksheetViewApisMutateExpectedSettings')),
                new TestCase('worksheet_protection_apis_mutate_expected_settings', array(Program::class, 'WorksheetProtectionApisMutateExpectedSettings')),
                new TestCase('autofilter_apis_mutate_expected_settings', array(Program::class, 'AutoFilterApisMutateExpectedSettings')),
                new TestCase('defined_name_apis_mutate_expected_settings', array(Program::class, 'DefinedNameApisMutateExpectedSettings')),
                new TestCase('hyperlink_apis_mutate_expected_settings', array(Program::class, 'HyperlinkApisMutateExpectedSettings')),
                new TestCase('validation_apis_mutate_expected_settings', array(Program::class, 'ValidationApisMutateExpectedSettings')),
                new TestCase('conditional_formatting_apis_mutate_expected_settings', array(Program::class, 'ConditionalFormattingApisMutateExpectedSettings')),
                new TestCase('conditional_formatting_advanced_apis_mutate_expected_settings', array(Program::class, 'ConditionalFormattingAdvancedApisMutateExpectedSettings')),
                new TestCase('page_setup_apis_mutate_expected_settings', array(Program::class, 'PageSetupApisMutateExpectedSettings')),
                new TestCase('workbook_metadata_apis_mutate_expected_settings', array(Program::class, 'WorkbookMetadataApisMutateExpectedSettings')),
                new TestCase('list_object_apis_mutate_expected_settings', array(Program::class, 'ListObjectApisMutateExpectedSettings')),
                new TestCase('picture_apis_mutate_expected_settings', array(Program::class, 'PictureApisMutateExpectedSettings')),
                new TestCase('shape_apis_mutate_expected_settings', array(Program::class, 'ShapeApisMutateExpectedSettings')),
                new TestCase('chart_collection_is_empty_on_new_worksheet', array(Program::class, 'ChartCollectionIsEmptyOnNewWorksheet')),
                new TestCase('chart_add_column_chart_returns_index_zero', array(Program::class, 'ChartAddColumnChartReturnsIndexZero')),
                new TestCase('chart_add_multiple_charts_increments_count', array(Program::class, 'ChartAddMultipleChartsIncrementsCount')),
                new TestCase('chart_add_sets_correct_position', array(Program::class, 'ChartAddSetsCorrectPosition')),
                new TestCase('chart_add_sets_correct_chart_type', array(Program::class, 'ChartAddSetsCorrectChartType')),
                new TestCase('chart_add_unsupported_chartex_type_throws', array(Program::class, 'ChartAddUnsupportedChartexTypeThrows')),
                new TestCase('chart_index_out_of_range_throws', array(Program::class, 'ChartIndexOutOfRangeThrows')),
                new TestCase('chart_roundtrip_column_preserves_count_and_type', array(Program::class, 'ChartRoundtripColumnPreservesCountAndType')),
                new TestCase('chart_roundtrip_bar_preserves_count_and_type', array(Program::class, 'ChartRoundtripBarPreservesCountAndType')),
                new TestCase('chart_roundtrip_line_preserves_count_and_type', array(Program::class, 'ChartRoundtripLinePreservesCountAndType')),
                new TestCase('chart_roundtrip_area_preserves_count_and_type', array(Program::class, 'ChartRoundtripAreaPreservesCountAndType')),
                new TestCase('chart_roundtrip_pie_preserves_count_and_type', array(Program::class, 'ChartRoundtripPiePreservesCountAndType')),
                new TestCase('chart_roundtrip_doughnut_preserves_count_and_type', array(Program::class, 'ChartRoundtripDoughnutPreservesCountAndType')),
                new TestCase('chart_roundtrip_scatter_preserves_count_and_type', array(Program::class, 'ChartRoundtripScatterPreservesCountAndType')),
                new TestCase('chart_roundtrip_radar_preserves_count_and_type', array(Program::class, 'ChartRoundtripRadarPreservesCountAndType')),
                new TestCase('chart_roundtrip_bubble_preserves_count_and_type', array(Program::class, 'ChartRoundtripBubblePreservesCountAndType')),
                new TestCase('chart_roundtrip_stock_preserves_count_and_type', array(Program::class, 'ChartRoundtripStockPreservesCountAndType')),
                new TestCase('chart_roundtrip_column3d_preserves_count_and_type', array(Program::class, 'ChartRoundtripColumn3DPreservesCountAndType')),
                new TestCase('chart_roundtrip_bar3d_preserves_count_and_type', array(Program::class, 'ChartRoundtripBar3DPreservesCountAndType')),
                new TestCase('chart_roundtrip_line3d_preserves_count_and_type', array(Program::class, 'ChartRoundtripLine3DPreservesCountAndType')),
                new TestCase('chart_roundtrip_area3d_preserves_count_and_type', array(Program::class, 'ChartRoundtripArea3DPreservesCountAndType')),
                new TestCase('chart_roundtrip_pie3d_preserves_count_and_type', array(Program::class, 'ChartRoundtripPie3DPreservesCountAndType')),
                new TestCase('chart_roundtrip_surface3d_preserves_count_and_type', array(Program::class, 'ChartRoundtripSurface3DPreservesCountAndType')),
                new TestCase('chart_roundtrip_contour_preserves_count_and_type', array(Program::class, 'ChartRoundtripContourPreservesCountAndType')),
                new TestCase('chart_worksheet_with_picture_and_chart_saves_both', array(Program::class, 'ChartWorksheetWithPictureAndChartSavesBoth')),
                new TestCase('comment_collection_empty_on_new_worksheet', array(Program::class, 'CommentCollectionEmptyOnNewWorksheet')),
                new TestCase('comment_add_by_row_col_sets_defaults', array(Program::class, 'CommentAddByRowColSetsDefaults')),
                new TestCase('comment_add_by_cell_name_parses_reference', array(Program::class, 'CommentAddByCellNameParsesReference')),
                new TestCase('comment_add_duplicate_cell_throws', array(Program::class, 'CommentAddDuplicateCellThrows')),
                new TestCase('comment_add_negative_row_throws', array(Program::class, 'CommentAddNegativeRowThrows')),
                new TestCase('comment_add_negative_column_throws', array(Program::class, 'CommentAddNegativeColumnThrows')),
                new TestCase('comment_add_invalid_cell_name_throws', array(Program::class, 'CommentAddInvalidCellNameThrows')),
                new TestCase('comment_count_increments_after_add', array(Program::class, 'CommentCountIncrementsAfterAdd')),
                new TestCase('comment_remove_at_index_decrements_count', array(Program::class, 'CommentRemoveAtIndexDecrementsCount')),
                new TestCase('comment_remove_at_index_out_of_range_throws', array(Program::class, 'CommentRemoveAtIndexOutOfRangeThrows')),
                new TestCase('comment_remove_at_cell_name_noop_when_not_found', array(Program::class, 'CommentRemoveAtCellNameNoopWhenNotFound')),
                new TestCase('comment_indexer_by_cell_name_returns_null_when_not_found', array(Program::class, 'CommentIndexerByCellNameReturnsNullWhenNotFound')),
                new TestCase('comment_note_and_author_mutate_via_property', array(Program::class, 'CommentNoteAndAuthorMutateViaProperty')),
                new TestCase('comment_is_visible_defaults_false', array(Program::class, 'CommentIsVisibleDefaultsFalse')),
                new TestCase('comment_width_height_defaults_are_129_and_75', array(Program::class, 'CommentWidthHeightDefaultsAre129And75')),
                new TestCase('comment_width_zero_throws', array(Program::class, 'CommentWidthZeroThrows')),
                new TestCase('comment_height_zero_throws', array(Program::class, 'CommentHeightZeroThrows')),
            )
        );
    }

    public static function A1IndexersRoundtrip()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells[[2, 27]]->PutValue('AB3');
        $sheet->Cells[[0, 0]]->PutValue(42);

        AssertEx::Equal('AB3', $sheet->Cells['AB3']->StringValue);
        AssertEx::Equal('42', $sheet->Cells[[0, 0]]->StringValue);
        AssertEx::Equal(CellValueType::IsString, $sheet->Cells['AB3']->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $sheet->Cells['A1']->Type);
    }

    public static function PutValueOverloadsAssignExpectedTypes()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $timestamp = new \DateTime('2024-05-06 07:08:09', new \DateTimeZone('UTC'));

        $sheet->Cells['A1']->PutValue('alpha');
        $sheet->Cells['B1']->PutValue(123);
        $sheet->Cells['C1']->PutValue(12.5);
        $sheet->Cells['D1']->PutValue(6.02214076E+23);
        $sheet->Cells['E1']->PutValue(true);
        $sheet->Cells['F1']->PutValue($timestamp);

        AssertEx::Equal(CellValueType::IsString, $sheet->Cells['A1']->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $sheet->Cells['B1']->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $sheet->Cells['C1']->Type);
        AssertEx::Equal(CellValueType::IsNumeric, $sheet->Cells['D1']->Type);
        AssertEx::Equal(CellValueType::IsBool, $sheet->Cells['E1']->Type);
        AssertEx::Equal(CellValueType::IsDateTime, $sheet->Cells['F1']->Type);

        AssertEx::Equal('alpha', $sheet->Cells['A1']->Value);
        AssertEx::Equal(123, (int) $sheet->Cells['B1']->Value);
        AssertEx::Equal(12.5, (float) $sheet->Cells['C1']->Value);
        AssertEx::True(abs((float) $sheet->Cells['D1']->Value - 6.02214076E+23) < 1E+10, 'Double PutValue should retain magnitude.');
        AssertEx::Equal(true, (bool) $sheet->Cells['E1']->Value);
        AssertEx::Equal($timestamp, $sheet->Cells['F1']->Value);
    }

    public static function StringValueFormatsSupportedScalarTypes()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $timestamp = new \DateTime('2024-05-06 07:08:09', new \DateTimeZone('UTC'));

        $sheet->Cells['A1']->PutValue(true);
        $sheet->Cells['B1']->PutValue(123);
        $sheet->Cells['C1']->PutValue(12.5);
        $sheet->Cells['D1']->PutValue($timestamp);

        AssertEx::Equal('TRUE', $sheet->Cells['A1']->StringValue);
        AssertEx::Equal('TRUE', $sheet->Cells['A1']->DisplayStringValue);
        AssertEx::Equal('123', $sheet->Cells['B1']->StringValue);
        AssertEx::Equal('123', $sheet->Cells['B1']->DisplayStringValue);
        AssertEx::Equal('12.5', $sheet->Cells['C1']->StringValue);
        AssertEx::Equal('12.5', $sheet->Cells['C1']->DisplayStringValue);
        AssertEx::Equal($timestamp->format('n/j/Y G:i'), $sheet->Cells['D1']->StringValue);
        AssertEx::Equal($timestamp->format('n/j/Y G:i'), $sheet->Cells['D1']->DisplayStringValue);
    }

    public static function CellsStylePropertyRoundtrip()
    {
        $workbook = new Workbook();
        $cells = $workbook->Worksheets[0]->Cells;

        $style = $cells->Style;
        $style->Font->Name = 'Tahoma';
        $style->Font->IsBold = true;
        $style->Number = 14;
        $cells->Style = $style;

        $persisted = $cells->Style;
        AssertEx::Equal('Tahoma', $persisted->Font->Name);
        AssertEx::True($persisted->Font->IsBold);
        AssertEx::Equal(14, $persisted->Number);
    }

    public static function PutValueStringAndObjectOverloadsFollowAsposeShape()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->Cells['A1']->PutValue('123', true);
        $sheet->Cells['B1']->PutValue('456', false, false);
        $sheet->Cells['C1']->PutValue(true);

        AssertEx::Equal(CellValueType::IsNumeric, $sheet->Cells['A1']->Type);
        AssertEx::Equal(123, (int) $sheet->Cells['A1']->Value);

        AssertEx::Equal(CellValueType::IsString, $sheet->Cells['B1']->Type);
        AssertEx::Equal('456', (string) $sheet->Cells['B1']->Value);

        AssertEx::Equal(CellValueType::IsBool, $sheet->Cells['C1']->Type);
        AssertEx::Equal(true, (bool) $sheet->Cells['C1']->Value);
    }

    public static function DisplayStringValueAppliesNumericFormatsAndStringValueStaysRaw()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $grouped = $sheet->Cells['A1'];
        $grouped->PutValue(1234.567);
        $groupedStyle = $grouped->GetStyle();
        $groupedStyle->NumberFormat = '#,##0.00';
        $grouped->SetStyle($groupedStyle);

        $percent = $sheet->Cells['B1'];
        $percent->PutValue(0.125);
        $percentStyle = $percent->GetStyle();
        $percentStyle->NumberFormat = '0.00%';
        $percent->SetStyle($percentStyle);

        $scientific = $sheet->Cells['C1'];
        $scientific->PutValue(1234);
        $scientificStyle = $scientific->GetStyle();
        $scientificStyle->NumberFormat = '0.00E+00';
        $scientific->SetStyle($scientificStyle);

        $fraction = $sheet->Cells['D1'];
        $fraction->PutValue(1.25);
        $fractionStyle = $fraction->GetStyle();
        $fractionStyle->NumberFormat = '# ?/?';
        $fraction->SetStyle($fractionStyle);

        $negative = $sheet->Cells['E1'];
        $negative->PutValue(-12.3);
        $negativeStyle = $negative->GetStyle();
        $negativeStyle->NumberFormat = '#,##0.00_);(#,##0.00)';
        $negative->SetStyle($negativeStyle);

        $color = $sheet->Cells['F1'];
        $color->PutValue(1.25);
        $colorStyle = $color->GetStyle();
        $colorStyle->NumberFormat = '[Blue]0.000';
        $color->SetStyle($colorStyle);

        $conditionalHigh = $sheet->Cells['G1'];
        $conditionalHigh->PutValue(125);
        $conditionalStyle = $conditionalHigh->GetStyle();
        $conditionalStyle->NumberFormat = '[>100]0.0;"small"';
        $conditionalHigh->SetStyle($conditionalStyle);

        $conditionalLow = $sheet->Cells['H1'];
        $conditionalLow->PutValue(10);
        $conditionalLow->SetStyle($conditionalStyle);

        AssertEx::Equal('1234.567', $grouped->StringValue);
        AssertEx::Equal('1,234.57', $grouped->DisplayStringValue);
        AssertEx::Equal('0.125', $percent->StringValue);
        AssertEx::Equal('12.50%', $percent->DisplayStringValue);
        AssertEx::Equal('1234', $scientific->StringValue);
        AssertEx::Equal('1.23E+03', $scientific->DisplayStringValue);
        AssertEx::Equal('1.25', $fraction->StringValue);
        AssertEx::Equal('1 1/4', $fraction->DisplayStringValue);
        AssertEx::Equal('-12.3', $negative->StringValue);
        AssertEx::Equal('(12.30)', $negative->DisplayStringValue);
        AssertEx::Equal('1.25', $color->StringValue);
        AssertEx::Equal('1.250', $color->DisplayStringValue);
        AssertEx::Equal('125', $conditionalHigh->StringValue);
        AssertEx::Equal('125.0', $conditionalHigh->DisplayStringValue);
        AssertEx::Equal('10', $conditionalLow->StringValue);
        AssertEx::Equal('small', $conditionalLow->DisplayStringValue);
    }

    public static function DisplayStringValueAppliesDateAndTextFormats()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $timestamp = new \DateTime('2024-05-06 07:08:09');

        $builtInDate = $sheet->Cells['A2'];
        $builtInDate->PutValue($timestamp);
        $builtInDateStyle = $builtInDate->GetStyle();
        $builtInDateStyle->Number = 14;
        $builtInDate->SetStyle($builtInDateStyle);

        $customDate = $sheet->Cells['B2'];
        $customDate->PutValue($timestamp);
        $customDateStyle = $customDate->GetStyle();
        $customDateStyle->NumberFormat = 'd-mmm-yy h:mm AM/PM';
        $customDate->SetStyle($customDateStyle);

        $elapsed = $sheet->Cells['C2'];
        $elapsed->PutValue($timestamp);
        $elapsedStyle = $elapsed->GetStyle();
        $elapsedStyle->Number = 46;
        $elapsed->SetStyle($elapsedStyle);

        $textCell = $sheet->Cells['D2'];
        $textCell->PutValue('ABC');
        $textStyle = $textCell->GetStyle();
        $textStyle->NumberFormat = '0;0;0;"Item "@';
        $textCell->SetStyle($textStyle);

        AssertEx::Equal('5/6/2024 7:08', $builtInDate->StringValue);
        AssertEx::Equal('05-06-24', $builtInDate->DisplayStringValue);
        AssertEx::Equal('5/6/2024 7:08', $customDate->StringValue);
        AssertEx::Equal('6-May-24 7:08 AM', $customDate->DisplayStringValue);
        AssertEx::Equal('5/6/2024 7:08', $elapsed->StringValue);
        AssertEx::Equal('1090039:08:09', $elapsed->DisplayStringValue);
        AssertEx::Equal('ABC', $textCell->StringValue);
        AssertEx::Equal('Item ABC', $textCell->DisplayStringValue);
    }

    public static function DisplayStringValueUsesWorkbookCultureAndLocaleDirectives()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $frCulture = CultureInfo::GetCultureInfo('fr-FR');
        $enCulture = CultureInfo::GetCultureInfo('en-US');
        $jpCulture = CultureInfo::GetCultureInfo('ja-JP');
        $deCulture = CultureInfo::GetCultureInfo('de-DE');
        $workbook->Settings->Culture = $frCulture;

        $numeric = $sheet->Cells['A3'];
        $numeric->PutValue(1234.5);
        $numericStyle = $numeric->GetStyle();
        $numericStyle->NumberFormat = '#,##0.00';
        $numeric->SetStyle($numericStyle);

        $dateCell = $sheet->Cells['B3'];
        $timestamp = new \DateTime('2024-05-06 07:08:09');
        $dateCell->PutValue($timestamp);
        $dateStyle = $dateCell->GetStyle();
        $dateStyle->NumberFormat = 'dddd, mmmm d, yyyy';
        $dateCell->SetStyle($dateStyle);

        $englishDate = $sheet->Cells['C3'];
        $englishDate->PutValue($timestamp);
        $englishDateStyle = $englishDate->GetStyle();
        $englishDateStyle->NumberFormat = '[$-409]dddd, mmmm d, yyyy';
        $englishDate->SetStyle($englishDateStyle);

        $yenCell = $sheet->Cells['D3'];
        $yenCell->PutValue(1234.5);
        $yenStyle = $yenCell->GetStyle();
        $yenStyle->NumberFormat = '[$' . "\u{00A5}" . '-411]#,##0.00';
        $yenCell->SetStyle($yenStyle);

        $longDateCell = $sheet->Cells['E3'];
        $longDateCell->PutValue($timestamp);
        $longDateStyle = $longDateCell->GetStyle();
        $longDateStyle->NumberFormat = '[$-F800]';
        $longDateCell->SetStyle($longDateStyle);

        AssertEx::Equal('1' . "\u{202F}" . '234,50', $numeric->DisplayStringValue);
        AssertEx::Equal('lundi, mai 6, 2024', $dateCell->DisplayStringValue);
        AssertEx::Equal('Monday, May 6, 2024', $englishDate->DisplayStringValue);
        AssertEx::Equal("\u{00A5}" . '1,234.50', $yenCell->DisplayStringValue);

        $workbook->Settings->Culture = $deCulture;
        AssertEx::Equal('Montag, 6. Mai 2024', $longDateCell->DisplayStringValue);
    }

    public static function DisplayStringValueUsesEnglishMonthAbbreviationsWithoutLocaleDirective()
    {
        $workbook = new Workbook();
        $workbook->Settings->Culture = CultureInfo::GetCultureInfo('zh-CN');
        $cell = $workbook->Worksheets[0]->Cells['A5'];
        $cell->PutValue(new \DateTime('2025-01-01 00:00:00'));
        $style = $cell->GetStyle();
        $style->NumberFormat = 'mmm-yy';
        $cell->SetStyle($style);

        AssertEx::Equal('Jan-25', $cell->DisplayStringValue);
    }

    public static function DisplayStringValueAppliesExtendedDateTokens()
    {
        $workbook = new Workbook();
        $workbook->Settings->Culture = CultureInfo::GetCultureInfo('en-US');
        $sheet = $workbook->Worksheets[0];
        $timestamp = new \DateTime('2024-05-06 07:08:09.345');

        $monthInitial = $sheet->Cells['A4'];
        $monthInitial->PutValue($timestamp);
        $monthInitialStyle = $monthInitial->GetStyle();
        $monthInitialStyle->NumberFormat = 'mmmmm d, yyyy';
        $monthInitial->SetStyle($monthInitialStyle);

        $abbreviatedDate = $sheet->Cells['B4'];
        $abbreviatedDate->PutValue($timestamp);
        $abbreviatedStyle = $abbreviatedDate->GetStyle();
        $abbreviatedStyle->NumberFormat = 'ddd, mmm d yyyy';
        $abbreviatedDate->SetStyle($abbreviatedStyle);

        $fractionalSeconds = $sheet->Cells['C4'];
        $fractionalSeconds->PutValue($timestamp);
        $fractionalStyle = $fractionalSeconds->GetStyle();
        $fractionalStyle->NumberFormat = 'h:mm:ss.000 AM/PM';
        $fractionalSeconds->SetStyle($fractionalStyle);

        $shortFraction = $sheet->Cells['D4'];
        $shortFraction->PutValue($timestamp);
        $shortFractionStyle = $shortFraction->GetStyle();
        $shortFractionStyle->NumberFormat = 'hh:mm:ss.00';
        $shortFraction->SetStyle($shortFractionStyle);

        $shortAmPm = $sheet->Cells['E4'];
        $shortAmPm->PutValue($timestamp);
        $shortAmPmStyle = $shortAmPm->GetStyle();
        $shortAmPmStyle->NumberFormat = 'h A/P';
        $shortAmPm->SetStyle($shortAmPmStyle);

        AssertEx::Equal('M 6, 2024', $monthInitial->DisplayStringValue);
        AssertEx::Equal('Mon, May 6 2024', $abbreviatedDate->DisplayStringValue);
        AssertEx::Equal('7:08:09.345 AM', $fractionalSeconds->DisplayStringValue);
        AssertEx::Equal('07:08:09.34', $shortFraction->DisplayStringValue);
        AssertEx::Equal('7 A', $shortAmPm->DisplayStringValue);
    }

    public static function DisplayStringValueAppliesLongTimeAndElapsedFractionWithCulture()
    {
        $workbook = new Workbook();
        $culture = CultureInfo::GetCultureInfo('de-DE');
        $workbook->Settings->Culture = $culture;
        $sheet = $workbook->Worksheets[0];
        $timestamp = new \DateTime('2024-05-06 07:08:09.345');

        $longTimeCell = $sheet->Cells['F4'];
        $longTimeCell->PutValue($timestamp);
        $longTimeStyle = $longTimeCell->GetStyle();
        $longTimeStyle->NumberFormat = '[$-F400]';
        $longTimeCell->SetStyle($longTimeStyle);

        $elapsedCell = $sheet->Cells['G4'];
        $elapsedCell->PutValue(new \DateTime('1899-12-31 07:08:09.345'));
        $elapsedStyle = $elapsedCell->GetStyle();
        $elapsedStyle->NumberFormat = '[h]:mm:ss.000';
        $elapsedCell->SetStyle($elapsedStyle);

        AssertEx::Equal('07:08:09', $longTimeCell->DisplayStringValue);
        AssertEx::Equal('7:08:09,345', $elapsedCell->DisplayStringValue);
    }

    public static function DisplayStringValueElapsedTimeHonors1904DateSystem()
    {
        $workbook = new Workbook();
        $workbook->Settings->Date1904 = true;
        $cell = $workbook->Worksheets[0]->Cells['A1'];
        $cell->PutValue(new \DateTime('1904-01-02 12:00:00'));
        $style = $cell->GetStyle();
        $style->NumberFormat = '[h]:mm:ss';
        $cell->SetStyle($style);

        AssertEx::Equal('36:00:00', $cell->DisplayStringValue);
    }

    public static function FormulaPropertyNormalizesAndPreservesCachedValue()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $cell = $sheet->Cells['C3'];

        $cell->PutValue(20);
        $cell->Formula = 'A1+B1';

        AssertEx::Equal(CellValueType::IsUnknown, $cell->Type);
        AssertEx::Equal('=A1+B1', $cell->Formula);
        AssertEx::Equal('20', $cell->StringValue);
        AssertEx::Equal(20, (int) $cell->Value);
    }

    public static function BlankCellsAreBlankByDefault()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['Z99'];

        AssertEx::Equal(CellValueType::IsNull, $cell->Type);
        AssertEx::Null($cell->Value);
        AssertEx::Equal('', $cell->StringValue);
    }

    public static function WorksheetNameAndCollectionGuards()
    {
        $workbook = new Workbook();
        $workbook->Worksheets->Add('Data');

        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->Worksheets->Add('data');
        });
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $unused = $workbook->Worksheets[0]->Cells['1A'];
        });
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $unused = $workbook->Worksheets[0]->Cells[[-1, 0]];
        });

        $workbook->Worksheets[0]->Cells['AB3']->PutValue('ok');
        AssertEx::Equal('ok', $workbook->Worksheets[0]->Cells[[2, 27]]->StringValue);
    }

    public static function StyleMutationRequiresSetStyleAndReturnsClones()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['A1'];

        $style = $cell->GetStyle();
        $style->Font->IsBold = true;
        $style->HorizontalAlignment = HorizontalAlignmentType::Right;

        $untouched = $cell->GetStyle();
        AssertEx::False($untouched->Font->IsBold);
        AssertEx::Equal(HorizontalAlignmentType::General, $untouched->HorizontalAlignment);

        $cell->SetStyle($style);
        $applied = $cell->GetStyle();
        AssertEx::True($applied->Font->IsBold);
        AssertEx::Equal(HorizontalAlignmentType::Right, $applied->HorizontalAlignment);

        $applied->Font->IsItalic = true;
        AssertEx::False($cell->GetStyle()->Font->IsItalic);
    }

    public static function StyleApiCoversAllPublicSettings()
    {
        $workbook = new Workbook();
        $primaryCell = $workbook->Worksheets[0]->Cells['A1'];
        $primaryCell->PutValue(1);

        $primaryStyle = $primaryCell->GetStyle();
        StyleScenarioFactory::ApplyPrimaryStyle($primaryStyle);

        $untouched = $primaryCell->GetStyle();
        AssertEx::Equal('Calibri', $untouched->Font->Name);
        AssertEx::Equal(FillPattern::None, $untouched->Pattern);
        AssertEx::True($untouched->IsLocked);
        AssertEx::False($untouched->IsHidden);
        AssertEx::Equal(0, $untouched->IndentLevel);
        AssertEx::Equal(0, $untouched->TextRotation);
        AssertEx::Equal(0, $untouched->ReadingOrder);
        AssertEx::False($untouched->ShrinkToFit);
        AssertEx::False($untouched->Font->IsStrikeout);
        AssertEx::False($untouched->Borders->DiagonalUp);
        AssertEx::False($untouched->Borders->DiagonalDown);
        AssertEx::Equal('General', $untouched->NumberFormat);

        $primaryCell->SetStyle($primaryStyle);
        StyleScenarioFactory::AssertPrimaryStyle($primaryCell->GetStyle());

        $mutatedClone = $primaryCell->GetStyle();
        $mutatedClone->Font->Name = 'Mutated';
        $mutatedClone->Font->IsStrikeout = false;
        $mutatedClone->Pattern = FillPattern::None;
        $mutatedClone->ForegroundColor = Color::Empty();
        $mutatedClone->BackgroundColor = Color::Empty();
        $mutatedClone->Borders->Right->LineStyle = BorderStyleType::None;
        $mutatedClone->Borders->Diagonal->LineStyle = BorderStyleType::None;
        $mutatedClone->Borders->DiagonalUp = false;
        $mutatedClone->Borders->DiagonalDown = false;
        $mutatedClone->Number = 0;
        $mutatedClone->HorizontalAlignment = HorizontalAlignmentType::Left;
        $mutatedClone->VerticalAlignment = VerticalAlignmentType::Bottom;
        $mutatedClone->WrapText = false;
        $mutatedClone->IndentLevel = 0;
        $mutatedClone->TextRotation = 0;
        $mutatedClone->ShrinkToFit = false;
        $mutatedClone->ReadingOrder = 0;
        $mutatedClone->RelativeIndent = 0;
        $mutatedClone->IsLocked = true;
        $mutatedClone->IsHidden = false;

        $persistedPrimary = $primaryCell->GetStyle();
        AssertEx::Equal('Arial', $persistedPrimary->Font->Name);
        AssertEx::True($persistedPrimary->Font->IsStrikeout);
        AssertEx::Equal(FillPattern::LightGrid, $persistedPrimary->Pattern);
        AssertEx::Equal(BorderStyleType::MediumDashDot, $persistedPrimary->Borders->Right->LineStyle);
        AssertEx::Equal(BorderStyleType::SlantedDashDot, $persistedPrimary->Borders->Diagonal->LineStyle);
        AssertEx::True($persistedPrimary->Borders->DiagonalUp);
        AssertEx::True($persistedPrimary->Borders->DiagonalDown);
        AssertEx::Equal(4, $persistedPrimary->Number);
        AssertEx::Equal('#,##0.00', $persistedPrimary->NumberFormat);
        AssertEx::Equal(HorizontalAlignmentType::Distributed, $persistedPrimary->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Distributed, $persistedPrimary->VerticalAlignment);
        AssertEx::True($persistedPrimary->WrapText);
        AssertEx::Equal(2, $persistedPrimary->IndentLevel);
        AssertEx::Equal(45, $persistedPrimary->TextRotation);
        AssertEx::True($persistedPrimary->ShrinkToFit);
        AssertEx::Equal(2, $persistedPrimary->ReadingOrder);
        AssertEx::Equal(1, $persistedPrimary->RelativeIndent);
        AssertEx::False($persistedPrimary->IsLocked);
        AssertEx::True($persistedPrimary->IsHidden);

        $customCell = $workbook->Worksheets[0]->Cells['B2'];
        $customStyle = $customCell->GetStyle();
        StyleScenarioFactory::ApplyCustomNumberStyle($customStyle);
        $customCell->SetStyle($customStyle);

        StyleScenarioFactory::AssertCustomNumberStyle($customCell->GetStyle());
        AssertEx::Equal(CellValueType::IsNull, $customCell->Type);

        $numberFormatStyle = new Style();
        $numberFormatStyle->NumberFormat = '0.00%';
        AssertEx::Equal(10, $numberFormatStyle->Number);
        AssertEx::Null($numberFormatStyle->Custom);
        $numberFormatStyle->NumberFormat = '[Blue]0.000';
        AssertEx::Equal(0, $numberFormatStyle->Number);
        AssertEx::Equal('[Blue]0.000', $numberFormatStyle->Custom);

        AssertEx::Throws(CellsException::class, function () use ($numberFormatStyle) {
            $numberFormatStyle->IndentLevel = -1;
        });
        AssertEx::Throws(CellsException::class, function () use ($numberFormatStyle) {
            $numberFormatStyle->TextRotation = 181;
        });
        AssertEx::Throws(CellsException::class, function () use ($numberFormatStyle) {
            $numberFormatStyle->ReadingOrder = 3;
        });
    }

    public static function FontEqualsAndToStringFollowAsposeShape()
    {
        $left = new Font();
        $left->Name = 'Arial';
        $left->Size = 12.0;
        $left->IsBold = true;
        $left->IsItalic = false;
        $left->Underline = FontUnderlineType::Single;
        $left->IsStrikeout = true;
        $left->Color = Color::FromArgb(255, 12, 34, 56);

        $right = new Font();
        $right->Name = 'Arial';
        $right->Size = 12.0;
        $right->IsBold = true;
        $right->IsItalic = false;
        $right->Underline = FontUnderlineType::Single;
        $right->IsStrikeout = true;
        $right->Color = Color::FromArgb(255, 12, 34, 56);

        AssertEx::True($left->Equals($right));

        $text = $left->ToString();
        AssertEx::True(strpos($text, 'Arial') !== false);
        AssertEx::True(strpos($text, 'Underline=Single') !== false);
    }

    public static function StyleEqualsHashCodeAreConsistent()
    {
        $left = new Style();
        $left->Font->Name = 'Arial';
        $left->Font->IsBold = true;
        $left->Borders->Left->LineStyle = BorderStyleType::Thin;
        $left->Borders->Left->Color = Color::FromArgb(255, 1, 2, 3);
        $left->Number = 14;
        $left->HorizontalAlignment = HorizontalAlignmentType::Center;
        $left->VerticalAlignment = VerticalAlignmentType::Top;
        $left->WrapText = true;
        $left->IsLocked = false;
        $left->IsHidden = true;

        $right = new Style();
        $right->Font->Name = 'Arial';
        $right->Font->IsBold = true;
        $right->Borders->Left->LineStyle = BorderStyleType::Thin;
        $right->Borders->Left->Color = Color::FromArgb(255, 1, 2, 3);
        $right->Number = 14;
        $right->HorizontalAlignment = HorizontalAlignmentType::Center;
        $right->VerticalAlignment = VerticalAlignmentType::Top;
        $right->WrapText = true;
        $right->IsLocked = false;
        $right->IsHidden = true;

        AssertEx::True($left->Equals($right));
        AssertEx::Equal($left->GetHashCode(), $right->GetHashCode());
    }

    public static function StyleCopyCopiesAllPublicSettings()
    {
        $source = new Style();
        $source->Font->Name = 'Consolas';
        $source->Font->IsBold = true;
        $source->Borders->Left->LineStyle = BorderStyleType::Medium;
        $source->Borders->Left->Color = Color::FromArgb(255, 9, 8, 7);
        $source->Pattern = FillPattern::Solid;
        $source->ForegroundColor = Color::FromArgb(255, 10, 20, 30);
        $source->BackgroundColor = Color::FromArgb(255, 40, 50, 60);
        $source->Number = 14;
        $source->Custom = '0.00';
        $source->HorizontalAlignment = HorizontalAlignmentType::Center;
        $source->VerticalAlignment = VerticalAlignmentType::Top;
        $source->WrapText = true;
        $source->IndentLevel = 2;
        $source->TextRotation = 45;
        $source->ShrinkToFit = true;
        $source->ReadingOrder = 2;
        $source->RelativeIndent = 1;
        $source->IsLocked = false;
        $source->IsHidden = true;

        $target = new Style();
        $target->Copy($source);

        AssertEx::True($target->Equals($source));

        $source->Font->Name = 'Mutated';
        AssertEx::Equal('Consolas', $target->Font->Name);
    }

    public static function CellGetStyleBoolOverloadMatchesGetStyle()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['A1'];
        $cell->PutValue(1);

        $style = $cell->GetStyle();
        $style->Font->Name = 'Calibri';
        $style->Font->IsBold = true;
        $cell->SetStyle($style);

        $fromDefault = $cell->GetStyle();
        $fromBoolFalse = $cell->GetStyle(false);
        $fromBoolTrue = $cell->GetStyle(true);

        AssertEx::True($fromDefault->Equals($fromBoolFalse));
        AssertEx::True($fromDefault->Equals($fromBoolTrue));
    }

    public static function CellSetStyleBoolOverloadMatchesSetStyle()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $left = $sheet->Cells['A1'];
        $right = $sheet->Cells['B1'];
        $left->PutValue(1);
        $right->PutValue(1);

        $style = $left->GetStyle();
        $style->Font->Name = 'Calibri';
        $style->Font->IsBold = true;
        $style->NumberFormat = '#,##0.00';

        $left->SetStyle($style);
        $right->SetStyle($style, true);

        AssertEx::True($left->GetStyle()->Equals($right->GetStyle()));
    }

    public static function CellSetStyleStyleFlagAppliesSelectedProperties()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['A1'];
        $cell->PutValue(1);

        $baseline = $cell->GetStyle();
        $baseline->Font->Name = 'Calibri';
        $baseline->Font->IsBold = false;
        $baseline->Font->Size = 11.0;
        $baseline->Font->IsItalic = false;
        $baseline->Font->Underline = FontUnderlineType::None;
        $baseline->Font->IsStrikeout = false;
        $baseline->Font->Color = Color::FromArgb(255, 0, 0, 0);
        $baseline->NumberFormat = '0';
        $baseline->HorizontalAlignment = HorizontalAlignmentType::Left;
        $baseline->IsHidden = false;
        $baseline->IsLocked = true;
        $baseline->QuotePrefix = false;
        $baseline->Borders->Left->LineStyle = BorderStyleType::None;
        $baseline->Borders->Right->LineStyle = BorderStyleType::None;
        $baseline->Borders->Top->LineStyle = BorderStyleType::None;
        $baseline->Borders->Bottom->LineStyle = BorderStyleType::None;
        $baseline->Borders->Diagonal->LineStyle = BorderStyleType::None;
        $baseline->Borders->DiagonalUp = false;
        $baseline->Borders->DiagonalDown = false;
        $cell->SetStyle($baseline);

        $incoming = $cell->GetStyle();
        $incoming->Font->Name = 'Arial';
        $incoming->Font->IsBold = true;
        $incoming->Font->Size = 14.0;
        $incoming->Font->IsItalic = true;
        $incoming->Font->Underline = FontUnderlineType::Single;
        $incoming->Font->IsStrikeout = true;
        $incoming->Font->Color = Color::FromArgb(255, 200, 10, 20);
        $incoming->NumberFormat = '#,##0.00';
        $incoming->HorizontalAlignment = HorizontalAlignmentType::Center;
        $incoming->IndentLevel = 3;
        $incoming->IsHidden = true;
        $incoming->IsLocked = false;
        $incoming->QuotePrefix = true;
        $incoming->Borders->Left->LineStyle = BorderStyleType::Thin;
        $incoming->Borders->Left->Color = Color::FromArgb(255, 1, 2, 3);
        $incoming->Borders->Right->LineStyle = BorderStyleType::Dashed;
        $incoming->Borders->Right->Color = Color::FromArgb(255, 30, 40, 50);
        $incoming->Borders->Top->LineStyle = BorderStyleType::Double;
        $incoming->Borders->Top->Color = Color::FromArgb(255, 60, 70, 80);
        $incoming->Borders->Bottom->LineStyle = BorderStyleType::Dotted;
        $incoming->Borders->Bottom->Color = Color::FromArgb(255, 90, 100, 110);
        $incoming->Borders->Diagonal->LineStyle = BorderStyleType::Medium;
        $incoming->Borders->Diagonal->Color = Color::FromArgb(255, 9, 8, 7);
        $incoming->Borders->DiagonalUp = true;
        $incoming->Borders->DiagonalDown = true;

        $flag = new StyleFlag();
        $flag->FontName = true;
        $flag->FontBold = true;
        $flag->FontSize = true;
        $flag->FontItalic = true;
        $flag->FontUnderline = true;
        $flag->FontStrike = true;
        $flag->FontColor = true;
        $flag->NumberFormat = true;
        $flag->Indent = true;
        $flag->Hidden = true;
        $flag->Locked = true;
        $flag->QuotePrefix = true;
        $flag->LeftBorder = true;
        $flag->RightBorder = true;
        $flag->TopBorder = true;
        $flag->BottomBorder = true;
        $flag->DiagonalUpBorder = true;
        $flag->DiagonalDownBorder = true;
        $cell->SetStyle($incoming, $flag);

        $applied = $cell->GetStyle();
        AssertEx::Equal('Arial', $applied->Font->Name);
        AssertEx::True($applied->Font->IsBold);
        AssertEx::Equal(14.0, $applied->Font->Size);
        AssertEx::True($applied->Font->IsItalic);
        AssertEx::Equal(FontUnderlineType::Single, $applied->Font->Underline);
        AssertEx::True($applied->Font->IsStrikeout);
        AssertEx::Equal(Color::FromArgb(255, 200, 10, 20), $applied->Font->Color);
        AssertEx::Equal('#,##0.00', $applied->NumberFormat);
        AssertEx::Equal(HorizontalAlignmentType::Left, $applied->HorizontalAlignment);
        AssertEx::Equal(3, $applied->IndentLevel);
        AssertEx::True($applied->IsHidden);
        AssertEx::False($applied->IsLocked);
        AssertEx::True($applied->QuotePrefix);
        AssertEx::Equal(BorderStyleType::Thin, $applied->Borders->Left->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 1, 2, 3), $applied->Borders->Left->Color);
        AssertEx::Equal(BorderStyleType::Dashed, $applied->Borders->Right->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 30, 40, 50), $applied->Borders->Right->Color);
        AssertEx::Equal(BorderStyleType::Double, $applied->Borders->Top->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 60, 70, 80), $applied->Borders->Top->Color);
        AssertEx::Equal(BorderStyleType::Dotted, $applied->Borders->Bottom->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 90, 100, 110), $applied->Borders->Bottom->Color);
        AssertEx::True($applied->Borders->DiagonalUp);
        AssertEx::True($applied->Borders->DiagonalDown);
        AssertEx::Equal(BorderStyleType::Medium, $applied->Borders->Diagonal->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 9, 8, 7), $applied->Borders->Diagonal->Color);
    }

    public static function CellSetStyleStyleFlagAllAppliesFullStyle()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['B2'];
        $cell->PutValue(1);

        $style = $cell->GetStyle();
        $style->Font->Name = 'Arial';
        $style->Font->IsBold = true;
        $style->NumberFormat = '#,##0.00';
        $style->HorizontalAlignment = HorizontalAlignmentType::Center;
        $style->VerticalAlignment = VerticalAlignmentType::Top;
        $style->WrapText = true;
        $style->IndentLevel = 2;
        $style->IsHidden = true;
        $style->IsLocked = false;
        $style->QuotePrefix = true;
        $style->Borders->Left->LineStyle = BorderStyleType::Thin;
        $style->Borders->Left->Color = Color::FromArgb(255, 10, 20, 30);

        $flag = new StyleFlag();
        $flag->All = true;
        $cell->SetStyle($style, $flag);

        $applied = $cell->GetStyle();
        AssertEx::True($applied->Equals($style));
    }

    public static function CellSetStyleStyleFlagAlignmentsAppliesIndentSettings()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['C3'];
        $cell->PutValue(1);

        $baseline = $cell->GetStyle();
        $baseline->HorizontalAlignment = HorizontalAlignmentType::Left;
        $baseline->VerticalAlignment = VerticalAlignmentType::Bottom;
        $baseline->WrapText = false;
        $baseline->IndentLevel = 0;
        $baseline->RelativeIndent = 0;
        $cell->SetStyle($baseline);

        $incoming = $cell->GetStyle();
        $incoming->HorizontalAlignment = HorizontalAlignmentType::Center;
        $incoming->VerticalAlignment = VerticalAlignmentType::Top;
        $incoming->WrapText = true;
        $incoming->IndentLevel = 4;
        $incoming->RelativeIndent = 2;

        $flag = new StyleFlag();
        $flag->Alignments = true;
        $cell->SetStyle($incoming, $flag);

        $applied = $cell->GetStyle();
        AssertEx::Equal(HorizontalAlignmentType::Center, $applied->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Top, $applied->VerticalAlignment);
        AssertEx::True($applied->WrapText);
        AssertEx::Equal(4, $applied->IndentLevel);
        AssertEx::Equal(2, $applied->RelativeIndent);
    }

    public static function WorksheetRowColumnAndMergeApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->VisibilityType = VisibilityType::Hidden;
        $sheet->Cells->Rows[2]->Height = 19.75;
        $sheet->Cells->Rows[4]->IsHidden = true;
        $sheet->Cells->Columns[1]->Width = 25.5;
        $sheet->Cells->Columns[3]->IsHidden = true;
        $sheet->Cells->Merge(1, 1, 2, 3);

        AssertEx::Equal(VisibilityType::Hidden, $sheet->VisibilityType);
        AssertEx::Equal(19.75, $sheet->Cells->Rows[2]->Height ?? 0);
        AssertEx::True($sheet->Cells->Rows[4]->IsHidden);
        AssertEx::Equal(25.5, $sheet->Cells->Columns[1]->Width ?? 0);
        AssertEx::True($sheet->Cells->Columns[3]->IsHidden);
        AssertEx::Equal(1, count($sheet->Cells->MergedCells));
        $merged = $sheet->Cells->MergedCells[0];
        AssertEx::Equal(1, $merged->StartRow);
        AssertEx::Equal(1, $merged->StartColumn);
        AssertEx::Equal(2, $merged->EndRow - $merged->StartRow + 1);
        AssertEx::Equal(3, $merged->EndColumn - $merged->StartColumn + 1);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Cells->Merge(2, 2, 2, 2);
        });
    }

    public static function WorksheetViewApisMutateExpectedSettings()
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

        $sheet->TabColor = Color::Empty();
        AssertEx::Equal(Color::Empty(), $sheet->TabColor);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Zoom = 9;
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Zoom = 401;
        });
    }

    public static function WorksheetProtectionApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->Protect();
        $sheet->Protection->Objects = true;
        $sheet->Protection->FormatCells = true;
        $sheet->Protection->InsertRows = true;
        $sheet->Protection->SelectUnlockedCells = true;

        AssertEx::True($sheet->Protection->IsProtected);
        AssertEx::True($sheet->Protection->Objects);
        AssertEx::True($sheet->Protection->FormatCells);
        AssertEx::True($sheet->Protection->InsertRows);
        AssertEx::True($sheet->Protection->SelectUnlockedCells);

        $sheet->Unprotect();
        AssertEx::False($sheet->Protection->IsProtected);
        AssertEx::False($sheet->Protection->Objects);
        AssertEx::False($sheet->Protection->FormatCells);
        AssertEx::False($sheet->Protection->InsertRows);
        AssertEx::False($sheet->Protection->SelectUnlockedCells);

        $sheet->Protection->AutoFilter = true;
        AssertEx::True($sheet->Protection->IsProtected);
        AssertEx::True($sheet->Protection->AutoFilter);
    }

    public static function AutoFilterApisMutateExpectedSettings()
    {
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        AutoFilterScenarioFactory::AssertAutoFilter($workbook);

        $sheet = $workbook->Worksheets[0];
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->AutoFilter->FilterColumns->Add(-1);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->AutoFilter->FilterColumns->Add(0);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->AutoFilter->SortState->SortConditions->Add('1A');
        });

        $sheet->AutoFilter->FilterColumns->RemoveAt(4);
        AssertEx::Equal(4, $sheet->AutoFilter->FilterColumns->Count);
        $sheet->AutoFilter->Clear();
        AssertEx::Equal('', $sheet->AutoFilter->Range);
        AssertEx::Equal(0, $sheet->AutoFilter->FilterColumns->Count);
        AssertEx::Equal(0, $sheet->AutoFilter->SortState->SortConditions->Count);
    }

    public static function DefinedNameApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Name = 'Data';
        $workbook->Worksheets->Add('Scoped');

        $total = $workbook->DefinedNames[$workbook->DefinedNames->Add('Total', '=SUM(Data!$A$1:$A$2)')];
        $total->Hidden = true;
        $total->Comment = 'Workbook scope';

        $scoped = $workbook->DefinedNames[$workbook->DefinedNames->Add('Input', "'Scoped'!\$B\$2", 1)];
        $scoped->Comment = 'Local scope';

        AssertEx::Equal(2, $workbook->DefinedNames->Count);
        AssertEx::Equal('Total', $total->Name);
        AssertEx::Equal('SUM(Data!$A$1:$A$2)', $total->Formula);
        AssertEx::Null($total->LocalSheetIndex);
        AssertEx::True($total->Hidden);
        AssertEx::Equal('Workbook scope', $total->Comment);

        AssertEx::Equal('Input', $scoped->Name);
        AssertEx::Equal("'Scoped'!\$B\$2", $scoped->Formula);
        AssertEx::Equal(1, $scoped->LocalSheetIndex ?? -1);
        AssertEx::Equal('Local scope', $scoped->Comment);

        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->DefinedNames->Add('Total', '1');
        });
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->DefinedNames->Add('_xlnm.Print_Area', 'A1');
        });
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->DefinedNames->Add('Broken', '1', 5);
        });

        $scoped->Name = 'Total';
        AssertEx::Throws(CellsException::class, function () use ($scoped) {
            $scoped->LocalSheetIndex = null;
        });

        $workbook->DefinedNames->RemoveAt(1);
        AssertEx::Equal(1, $workbook->DefinedNames->Count);
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->DefinedNames->RemoveAt(5);
        });
    }

    public static function HyperlinkApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->Cells['A1']->PutValue('Docs');
        $external = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('A1', 1, 1, 'https://example.com/docs')];
        $external->TextToDisplay = 'Docs';
        $external->ScreenTip = 'External docs';

        $internalLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('B2', 1, 1, 'Sheet1!C3')];
        $internalLink->TextToDisplay = 'Jump';

        $rangeLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('C4', 2, 2, 'mailto:test@example.com')];
        $rangeLink->ScreenTip = 'Send mail';

        $defaultTextLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('E7', 1, 1, 'https://default-text.example.com')];

        AssertEx::Equal(4, $sheet->Hyperlinks->Count);
        AssertEx::Equal('A1', $external->Area);
        AssertEx::Equal('https://example.com/docs', $external->Address);
        AssertEx::Equal('External docs', $external->ScreenTip);
        AssertEx::Equal('Docs', $external->TextToDisplay);

        AssertEx::Equal('B2', $internalLink->Area);
        AssertEx::Equal('Sheet1!C3', $internalLink->Address);
        AssertEx::Equal('Jump', $internalLink->TextToDisplay);

        AssertEx::Equal('C4:D5', $rangeLink->Area);
        AssertEx::Equal('mailto:test@example.com', $rangeLink->Address);
        AssertEx::Equal('Send mail', $rangeLink->ScreenTip);
        AssertEx::Equal('https://default-text.example.com', $defaultTextLink->TextToDisplay);

        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Hyperlinks->Add('A1', 1, 1, 'https://overlap.example.com');
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Hyperlinks->Add('Z1', 0, 1, 'https://invalid.example.com');
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Hyperlinks->Add('A2', 1, 1, '');
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Hyperlinks[-1];
        });

        $sheet->Hyperlinks->RemoveAt(1);
        AssertEx::Equal(3, $sheet->Hyperlinks->Count);
        AssertEx::Equal('C4:D5', $sheet->Hyperlinks[1]->Area);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Hyperlinks->RemoveAt(99);
        });

        $replacementIndex = $sheet->Hyperlinks->Add('F1', 1, 1, 'https://replace.example.com');
        $replacement = $sheet->Hyperlinks[$replacementIndex];
        $replacement->TextToDisplay = 'Replacement';
        $sheet->Hyperlinks->RemoveAt($replacementIndex);
        $sheet->Hyperlinks[0] = $replacement;
        AssertEx::Equal('F1', $sheet->Hyperlinks[0]->Area);
        AssertEx::Equal('https://replace.example.com', $sheet->Hyperlinks[0]->Address);
        AssertEx::Equal('Replacement', $sheet->Hyperlinks[0]->TextToDisplay);
        AssertEx::Throws(\InvalidArgumentException::class, function () use ($sheet) {
            $sheet->Hyperlinks[0] = null;
        });

        $currentCapacity = $sheet->Hyperlinks->Capacity;
        $sheet->Hyperlinks->Capacity = $currentCapacity + 4;
        AssertEx::True($sheet->Hyperlinks->Capacity >= $currentCapacity + 4);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Hyperlinks->Capacity = 1;
        });

        $sheet->Hyperlinks->Clear();
        AssertEx::Equal(0, $sheet->Hyperlinks->Count);
    }

    public static function ValidationApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $primary = $sheet->Validations[$sheet->Validations->Add(CellArea::CreateCellArea('A1', 'B2'))];
        $primary->Type = ValidationType::WholeNumber;
        $primary->Operator = OperatorType::Between;
        $primary->Formula1 = '=1';
        $primary->Formula2 = '=10';
        $primary->ShowError = true;
        $primary->ErrorTitle = 'Whole Number';
        $primary->ErrorMessage = 'Enter 1-10';
        $primary->AddArea(CellArea::CreateCellArea('D4', 'D5'));

        AssertEx::Equal(1, $sheet->Validations->Count);
        AssertEx::Equal(2, count($primary->Areas));
        AssertEx::Equal('1', $primary->Formula1);
        AssertEx::Equal('10', $primary->Formula2);
        AssertEx::Equal(ValidationType::WholeNumber, $sheet->Validations->GetValidationInCell(0, 0)->Type);
        AssertEx::Equal(ValidationType::WholeNumber, $sheet->Validations->GetValidationInCell(4, 3)->Type);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Validations->Add(CellArea::CreateCellArea('B2', 'C3'));
        });

        $secondIndex = $sheet->Validations->Add(CellArea::CreateCellArea('F1', 'F1'));
        $second = $sheet->Validations[$secondIndex];
        $second->Type = ValidationType::List;
        $second->Formula1 = '"Y,N"';
        AssertEx::Equal(2, $sheet->Validations->Count);

        $sheet->Validations->RemoveACell(0, 0);
        AssertEx::Null($sheet->Validations->GetValidationInCell(0, 0));
        AssertEx::NotNull($sheet->Validations->GetValidationInCell(0, 1));

        $sheet->Validations->RemoveArea(CellArea::CreateCellArea('F1', 'F1'));
        AssertEx::Equal(1, $sheet->Validations->Count);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Validations->RemoveACell(-1, 0);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Validations[-1];
        });
    }

    public static function ConditionalFormattingApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $index = $sheet->ConditionalFormattings->Add();
        $collection = $sheet->ConditionalFormattings[$index];
        $collection->AddArea(CellArea::CreateCellArea('A1', 'A3'));
        $conditionIndex = $collection->AddCondition(FormatConditionType::CellValue, OperatorType::Between, '=1', '=9');
        $condition = $collection[$conditionIndex];
        $condition->StopIfTrue = true;
        $condition->Priority = 1;
        $style = $condition->Style;
        $style->Pattern = FillPattern::Solid;
        $style->ForegroundColor = Color::FromArgb(255, 255, 0, 0);
        $condition->Style = $style;

        AssertEx::Equal(1, $sheet->ConditionalFormattings->Count);
        AssertEx::Equal(1, $collection->RangeCount);
        AssertEx::Equal(1, $collection->Count);
        AssertEx::Equal('1', $condition->Formula1);
        AssertEx::Equal('9', $condition->Formula2);
        AssertEx::True($condition->StopIfTrue);
        AssertEx::Equal(FillPattern::Solid, $condition->Style->Pattern);

        $collection->AddCondition(FormatConditionType::Expression);
        AssertEx::Equal(2, $collection->Count);
        $collection->RemoveCondition(1);
        AssertEx::Equal(1, $collection->Count);

        $collection->AddArea(CellArea::CreateCellArea('C1', 'C2'));
        AssertEx::Equal(2, $collection->RangeCount);
        $collection->RemoveArea(0, 0, 1, 1);
        AssertEx::Equal(2, $collection->RangeCount);
        AssertEx::Equal(0, $collection->GetCellArea(0)->StartRow);
        AssertEx::Equal(2, $collection->GetCellArea(0)->StartColumn);

        $sheet->ConditionalFormattings->RemoveArea(0, 2, 2, 1);
        AssertEx::Equal(1, $sheet->ConditionalFormattings->Count);
        $sheet->ConditionalFormattings->RemoveAt(0);
        AssertEx::Equal(0, $sheet->ConditionalFormattings->Count);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->ConditionalFormattings[-1];
        });
    }

    public static function ConditionalFormattingAdvancedApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $contains = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $contains->AddArea(CellArea::CreateCellArea('A1', 'A10'));
        $containsRule = $contains[$contains->AddCondition(FormatConditionType::ContainsText)];
        $containsRule->Formula1 = 'error';
        $containsRule->Priority = 2;

        $timePeriod = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $timePeriod->AddArea(CellArea::CreateCellArea('B1', 'B10'));
        $timeRule = $timePeriod[$timePeriod->AddCondition(FormatConditionType::TimePeriod)];
        $timeRule->TimePeriod = 'today';

        $top10 = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $top10->AddArea(CellArea::CreateCellArea('C1', 'C10'));
        $topRule = $top10[$top10->AddCondition(FormatConditionType::Top10)];
        $topRule->Percent = true;
        $topRule->Rank = 10;

        $colorScale = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $colorScale->AddArea(CellArea::CreateCellArea('D1', 'D10'));
        $colorRule = $colorScale[$colorScale->AddCondition(FormatConditionType::ColorScale)];
        $colorRule->ColorScaleCount = 3;
        $colorRule->MinColor = Color::FromArgb(255, 248, 105, 107);
        $colorRule->MidColor = Color::FromArgb(255, 255, 235, 132);
        $colorRule->MaxColor = Color::FromArgb(255, 99, 190, 123);

        $dataBar = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $dataBar->AddArea(CellArea::CreateCellArea('E1', 'E10'));
        $dataBarRule = $dataBar[$dataBar->AddCondition(FormatConditionType::DataBar)];
        $dataBarRule->BarColor = Color::FromArgb(255, 99, 142, 198);
        $dataBarRule->ShowBorder = true;
        $dataBarRule->Direction = 'left-to-right';

        $iconSet = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $iconSet->AddArea(CellArea::CreateCellArea('F1', 'F10'));
        $iconSetRule = $iconSet[$iconSet->AddCondition(FormatConditionType::IconSet)];
        $iconSetRule->IconSetType = '4Arrows';
        $iconSetRule->ReverseIcons = true;
        $iconSetRule->ShowIconOnly = true;

        AssertEx::Equal(6, $sheet->ConditionalFormattings->Count);
        AssertEx::Equal(FormatConditionType::ContainsText, $containsRule->Type);
        AssertEx::Equal('error', $containsRule->Formula1);
        AssertEx::Equal(2, $containsRule->Priority);
        AssertEx::Equal('today', $timeRule->TimePeriod);
        AssertEx::True($topRule->Percent);
        AssertEx::Equal(10, $topRule->Rank);
        AssertEx::Equal(3, $colorRule->ColorScaleCount);
        AssertEx::Equal(Color::FromArgb(255, 248, 105, 107), $colorRule->MinColor);
        AssertEx::Equal(Color::FromArgb(255, 99, 142, 198), $dataBarRule->BarColor);
        AssertEx::True($dataBarRule->ShowBorder);
        AssertEx::Equal('4Arrows', $iconSetRule->IconSetType);
        AssertEx::True($iconSetRule->ReverseIcons);
        AssertEx::True($iconSetRule->ShowIconOnly);
    }

    public static function PageSetupApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $pageSetup = $workbook->Worksheets[0]->PageSetup;

        $pageSetup->LeftMargin = 0.508;
        $pageSetup->RightMargin = 0.635;
        $pageSetup->Orientation = PageOrientationType::Landscape;
        $pageSetup->PaperSize = PaperSizeType::PaperA4;
        $pageSetup->FirstPageNumber = 2;
        $pageSetup->Scale = 90;
        $pageSetup->FitToPagesWide = 1;
        $pageSetup->FitToPagesTall = 3;
        $pageSetup->PrintArea = '$A$1:$D$20';
        $pageSetup->PrintTitleRows = '$1:$2';
        $pageSetup->PrintTitleColumns = '$A:$B';
        $pageSetup->LeftHeader = 'LH';
        $pageSetup->CenterFooter = 'CF';
        $pageSetup->PrintGridlines = true;
        $pageSetup->CenterHorizontally = true;
        $pageSetup->AddHorizontalPageBreak(5);
        $pageSetup->AddVerticalPageBreak(2);

        AssertEx::Equal(0.508, $pageSetup->LeftMargin);
        AssertEx::Equal(0.635, $pageSetup->RightMargin);
        AssertEx::Equal(PageOrientationType::Landscape, $pageSetup->Orientation);
        AssertEx::Equal(PaperSizeType::PaperA4, $pageSetup->PaperSize);
        AssertEx::Equal(2, $pageSetup->FirstPageNumber ?? 0);
        AssertEx::Equal(90, $pageSetup->Scale ?? 0);
        AssertEx::Equal(1, $pageSetup->FitToPagesWide ?? 0);
        AssertEx::Equal(3, $pageSetup->FitToPagesTall ?? 0);
        AssertEx::Equal('$A$1:$D$20', $pageSetup->PrintArea);
        AssertEx::Equal('$1:$2', $pageSetup->PrintTitleRows);
        AssertEx::Equal('$A:$B', $pageSetup->PrintTitleColumns);
        AssertEx::Equal('LH', $pageSetup->LeftHeader);
        AssertEx::Equal('CF', $pageSetup->CenterFooter);
        AssertEx::True($pageSetup->PrintGridlines);
        AssertEx::True($pageSetup->CenterHorizontally);
        AssertEx::Equal(1, count($pageSetup->HorizontalPageBreaks));
        AssertEx::Equal(5, $pageSetup->HorizontalPageBreaks[0]);
        AssertEx::Equal(1, count($pageSetup->VerticalPageBreaks));
        AssertEx::Equal(2, $pageSetup->VerticalPageBreaks[0]);
        AssertEx::Throws(CellsException::class, function () use ($pageSetup) {
            $pageSetup->Scale = 5;
        });
        AssertEx::Throws(CellsException::class, function () use ($pageSetup) {
            $pageSetup->LeftMargin = -0.1;
        });
    }

    public static function WorkbookMetadataApisMutateExpectedSettings()
    {
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($workbook);
    }

    public static function ListObjectApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->Cells['A1']->PutValue('Name');
        $sheet->Cells['B1']->PutValue('Score');
        $sheet->Cells['A2']->PutValue('Alice');
        $sheet->Cells['B2']->PutValue(95);
        $sheet->Cells['A3']->PutValue('Bob');
        $sheet->Cells['B3']->PutValue(87);

        $index = $sheet->ListObjects->Add('A1', 'B3', true);
        AssertEx::Equal(0, $index);
        AssertEx::Equal(1, $sheet->ListObjects->Count);

        $table = $sheet->ListObjects[$index];
        AssertEx::Equal('Table1', $table->DisplayName);
        AssertEx::Equal(0, $table->StartRow);
        AssertEx::Equal(0, $table->StartColumn);
        AssertEx::Equal(2, $table->EndRow);
        AssertEx::Equal(1, $table->EndColumn);
        AssertEx::True($table->ShowHeaderRow);
        AssertEx::False($table->ShowTotals);
        AssertEx::Equal(2, $table->ListColumns->Count);
        AssertEx::Equal('Name', $table->ListColumns[0]->Name);
        AssertEx::Equal('Score', $table->ListColumns[1]->Name);

        $table->TableStyleType = TableStyleType::TableStyleMedium9;
        AssertEx::Equal(TableStyleType::TableStyleMedium9, $table->TableStyleType);
        AssertEx::Equal('TableStyleMedium9', $table->TableStyleName);

        $table->TableStyleName = 'MyCustom';
        AssertEx::Equal('MyCustom', $table->TableStyleName);
        AssertEx::Equal(TableStyleType::Custom, $table->TableStyleType);

        $table->TableStyleType = TableStyleType::None;
        AssertEx::Equal('', $table->TableStyleName);

        $table->ShowTableStyleRowStripes = true;
        $table->ShowTableStyleFirstColumn = true;
        $table->ShowTableStyleLastColumn = false;
        $table->ShowTableStyleColumnStripes = false;
        AssertEx::True($table->ShowTableStyleRowStripes);
        AssertEx::True($table->ShowTableStyleFirstColumn);
        AssertEx::False($table->ShowTableStyleLastColumn);
        AssertEx::False($table->ShowTableStyleColumnStripes);

        $table->ShowTotals = true;
        $scoreCol = $table->ListColumns[1];
        $scoreCol->TotalsCalculation = TotalsCalculation::Sum;
        $scoreCol->TotalsRowLabel = '';
        AssertEx::Equal(TotalsCalculation::Sum, $table->ListColumns[1]->TotalsCalculation);

        $nameCol = $table->ListColumns[0];
        $nameCol->TotalsCalculation = TotalsCalculation::None;
        $nameCol->TotalsRowLabel = 'Total';
        AssertEx::Equal(TotalsCalculation::None, $table->ListColumns[0]->TotalsCalculation);
        AssertEx::Equal('Total', $table->ListColumns[0]->TotalsRowLabel);

        $scoreCol->TotalsCalculation = TotalsCalculation::Average;
        AssertEx::Equal(TotalsCalculation::Average, $table->ListColumns[1]->TotalsCalculation);

        $table->DisplayName = 'SalesTable';
        AssertEx::Equal('SalesTable', $table->DisplayName);
        AssertEx::Equal('SalesTable', $sheet->ListObjects['SalesTable']->DisplayName);

        $table->Comment = 'Test comment';
        AssertEx::Equal('Test comment', $table->Comment);

        $table->ShowAutoFilter();
        $table->RemoveAutoFilter();

        $table->Resize(0, 0, 3, 1, true);
        AssertEx::Equal(3, $table->EndRow);

        AssertEx::Throws(CellsException::class, function () use ($table) {
            $table->DisplayName = '';
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->ListObjects->Add(0, 0, 2, 1, true);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->ListObjects[5];
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->ListObjects->RemoveAt(5);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->ListObjects['NoSuchTable'];
        });

        $sheet->Cells['D1']->PutValue('Category');
        $sheet->Cells['E1']->PutValue('Units');
        $sheet->Cells['D2']->PutValue('A');
        $sheet->Cells['E2']->PutValue(10);
        $index2 = $sheet->ListObjects->Add(0, 3, 2, 4, true);
        $table2 = $sheet->ListObjects[$index2];
        AssertEx::Equal(2, $sheet->ListObjects->Count);
        AssertEx::Equal('Table2', $table2->DisplayName);
        AssertEx::Equal('Category', $table2->ListColumns[0]->Name);
        AssertEx::Equal('Units', $table2->ListColumns[1]->Name);

        $table2->ConvertToRange();
        AssertEx::Equal(1, $sheet->ListObjects->Count);

        $sheet->ListObjects->RemoveAt(0);
        AssertEx::Equal(0, $sheet->ListObjects->Count);
    }

    public static function PictureApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $jpegBytes = PictureScenarioFactory::LoadJpeg();
        $pngBytes = PictureScenarioFactory::LoadPng();

        AssertEx::Equal(0, $sheet->Pictures->Count);

        $idx = $sheet->Pictures->Add(1, 2, 5, 6, $jpegBytes);
        AssertEx::Equal(0, $idx);
        AssertEx::Equal(1, $sheet->Pictures->Count);

        $pic = $sheet->Pictures[0];
        AssertEx::Equal('Picture 1', $pic->Name);
        AssertEx::Equal(1, $pic->UpperLeftRow);
        AssertEx::Equal(2, $pic->UpperLeftColumn);
        AssertEx::Equal(5, $pic->LowerRightRow);
        AssertEx::Equal(6, $pic->LowerRightColumn);
        AssertEx::Equal(ImageType::Jpeg, $pic->ImageType);
        AssertEx::Equal(strlen($jpegBytes), strlen($pic->Data));

        $pic->Name = 'Logo';
        AssertEx::Equal('Logo', $sheet->Pictures[0]->Name);

        $pic->UpperLeftRow = 0;
        $pic->UpperLeftColumn = 0;
        $pic->LowerRightRow = 3;
        $pic->LowerRightColumn = 3;
        AssertEx::Equal(0, $sheet->Pictures[0]->UpperLeftRow);
        AssertEx::Equal(3, $sheet->Pictures[0]->LowerRightRow);

        $idx2 = $sheet->Pictures->Add(5, 0, 8, 4, $pngBytes);
        AssertEx::Equal(1, $idx2);
        AssertEx::Equal(2, $sheet->Pictures->Count);
        AssertEx::Equal(ImageType::Png, $sheet->Pictures[1]->ImageType);

        $sheet->Pictures->RemoveAt(0);
        AssertEx::Equal(1, $sheet->Pictures->Count);
        AssertEx::Equal(ImageType::Png, $sheet->Pictures[0]->ImageType);

        AssertEx::Throws(CellsException::class, function () use ($sheet, $jpegBytes) {
            $sheet->Pictures->Add(-1, 0, 3, 3, $jpegBytes);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Pictures->Add(0, 0, 3, 3, null);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Pictures->Add(0, 0, 3, 3, '');
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Pictures->RemoveAt(99);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Pictures[99];
        });
    }

    public static function ShapeApisMutateExpectedSettings()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        AssertEx::Equal(0, $sheet->Shapes->Count);

        $idx = $sheet->Shapes->Add(1, 1, 5, 3, AutoShapeType::Rectangle);
        AssertEx::Equal(0, $idx);
        AssertEx::Equal(1, $sheet->Shapes->Count);

        $shape = $sheet->Shapes[0];
        AssertEx::Equal('Shape 1', $shape->Name);
        AssertEx::Equal(1, $shape->UpperLeftRow);
        AssertEx::Equal(1, $shape->UpperLeftColumn);
        AssertEx::Equal(5, $shape->LowerRightRow);
        AssertEx::Equal(3, $shape->LowerRightColumn);
        AssertEx::Equal(AutoShapeType::Rectangle, $shape->AutoShapeType);
        AssertEx::Equal('rect', $shape->GeometryType);

        $shape->Name = 'MyRect';
        AssertEx::Equal('MyRect', $sheet->Shapes[0]->Name);

        $shape->UpperLeftRow = 0;
        $shape->UpperLeftColumn = 0;
        $shape->LowerRightRow = 4;
        $shape->LowerRightColumn = 2;
        AssertEx::Equal(0, $sheet->Shapes[0]->UpperLeftRow);
        AssertEx::Equal(4, $sheet->Shapes[0]->LowerRightRow);

        $shape->AutoShapeType = AutoShapeType::RightArrow;
        AssertEx::Equal(AutoShapeType::RightArrow, $sheet->Shapes[0]->AutoShapeType);
        AssertEx::Equal('rightArrow', $sheet->Shapes[0]->GeometryType);

        $shape->GeometryType = 'star12';
        AssertEx::Equal(AutoShapeType::Star12Point, $sheet->Shapes[0]->AutoShapeType);

        $idx2 = $sheet->Shapes->Add(7, 1, 9, 3, AutoShapeType::Ellipse);
        AssertEx::Equal(1, $idx2);
        AssertEx::Equal(2, $sheet->Shapes->Count);
        AssertEx::Equal(AutoShapeType::Ellipse, $sheet->Shapes[1]->AutoShapeType);

        $sheet->Shapes->RemoveAt(0);
        AssertEx::Equal(1, $sheet->Shapes->Count);
        AssertEx::Equal(AutoShapeType::Ellipse, $sheet->Shapes[0]->AutoShapeType);

        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Shapes->Add(-1, 0, 3, 3, AutoShapeType::Rectangle);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Shapes->Add(0, 2, 3, 1, AutoShapeType::Rectangle);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Shapes->RemoveAt(99);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Shapes[99];
        });
    }

    public static function ChartCollectionIsEmptyOnNewWorksheet()
    {
        $workbook = new Workbook();
        AssertEx::Equal(0, $workbook->Worksheets[0]->Charts->Count);
    }

    public static function ChartAddColumnChartReturnsIndexZero()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $idx = $sheet->Charts->Add(ChartType::Column, 'Sheet1!$B$1:$B$5', 5, 0, 15, 8);
        AssertEx::Equal(0, $idx);
        AssertEx::Equal(1, $sheet->Charts->Count);
    }

    public static function ChartAddMultipleChartsIncrementsCount()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Charts->Add(ChartType::Column, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        $sheet->Charts->Add(ChartType::Line, 'Sheet1!$C$1:$C$5', 12, 0, 22, 8);
        AssertEx::Equal(2, $sheet->Charts->Count);
        AssertEx::Equal(ChartType::Column, $sheet->Charts[0]->ChartType);
        AssertEx::Equal(ChartType::Line, $sheet->Charts[1]->ChartType);
    }

    public static function ChartAddSetsCorrectPosition()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Charts->Add(ChartType::Bar, 'Sheet1!$B$1:$B$5', 7, 2, 20, 12);
        $chart = $sheet->Charts[0];
        AssertEx::Equal(7, $chart->UpperLeftRow);
        AssertEx::Equal(2, $chart->UpperLeftColumn);
        AssertEx::Equal(20, $chart->LowerRightRow);
        AssertEx::Equal(12, $chart->LowerRightColumn);
    }

    public static function ChartAddSetsCorrectChartType()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Charts->Add(ChartType::Pie, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        AssertEx::Equal(ChartType::Pie, $sheet->Charts[0]->ChartType);
    }

    public static function ChartAddUnsupportedChartexTypeThrows()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        AssertEx::Throws(UnsupportedFeatureException::class, function () use ($sheet) {
            $sheet->Charts->Add(ChartType::Waterfall, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        });
        AssertEx::Throws(UnsupportedFeatureException::class, function () use ($sheet) {
            $sheet->Charts->Add(ChartType::Treemap, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        });
        AssertEx::Throws(UnsupportedFeatureException::class, function () use ($sheet) {
            $sheet->Charts->Add(ChartType::Funnel, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        });
    }

    public static function ChartIndexOutOfRangeThrows()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Charts[0];
        });
        $sheet->Charts->Add(ChartType::Column, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Charts[1];
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $unused = $sheet->Charts[-1];
        });
    }

    public static function ChartRoundtripColumnPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Column);
    }

    public static function ChartRoundtripBarPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Bar);
    }

    public static function ChartRoundtripLinePreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Line);
    }

    public static function ChartRoundtripAreaPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Area);
    }

    public static function ChartRoundtripPiePreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Pie);
    }

    public static function ChartRoundtripDoughnutPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Doughnut);
    }

    public static function ChartRoundtripScatterPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Scatter);
    }

    public static function ChartRoundtripRadarPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Radar);
    }

    public static function ChartRoundtripBubblePreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Bubble);
    }

    public static function ChartRoundtripStockPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Stock);
    }

    public static function ChartRoundtripColumn3DPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Column3D);
    }

    public static function ChartRoundtripBar3DPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Bar3D);
    }

    public static function ChartRoundtripLine3DPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Line3D);
    }

    public static function ChartRoundtripArea3DPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Area3D);
    }

    public static function ChartRoundtripPie3DPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Pie3D);
    }

    public static function ChartRoundtripSurface3DPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Surface3D);
    }

    public static function ChartRoundtripContourPreservesCountAndType()
    {
        self::RoundtripChart(ChartType::Contour);
    }

    private static function RoundtripChart($type)
    {
        $temp = new TemporaryDirectory('chart-unit-roundtrips');
        $outputPath = $temp->GetPath($type . '.xlsx');

        $workbook = ChartScenarioFactory::CreateChartWorkbook($type);
        $workbook->Save($outputPath);

        $reloaded = new Workbook($outputPath);
        ChartScenarioFactory::AssertChartWorkbook($reloaded, $type);
    }

    public static function ChartWorksheetWithPictureAndChartSavesBoth()
    {
        $jpegBytes = PictureScenarioFactory::LoadJpeg();
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Pictures->Add(0, 0, 5, 5, $jpegBytes);
        $sheet->Charts->Add(ChartType::Column, 'Sheet1!$B$1:$B$5', 6, 0, 20, 10);
        AssertEx::Equal(1, $sheet->Pictures->Count);
        AssertEx::Equal(1, $sheet->Charts->Count);

        $temp = new TemporaryDirectory('chart-picture-mixed');
        $path = $temp->GetPath('chart-picture.xlsx');
        $workbook->Save($path);
        $reloaded = new Workbook($path);
        AssertEx::Equal(1, $reloaded->Worksheets[0]->Pictures->Count);
        AssertEx::Equal(1, $reloaded->Worksheets[0]->Charts->Count);
        AssertEx::Equal(ChartType::Column, $reloaded->Worksheets[0]->Charts[0]->ChartType);
    }

    public static function CommentCollectionEmptyOnNewWorksheet()
    {
        $workbook = new Workbook();
        AssertEx::Equal(0, $workbook->Worksheets[0]->Comments->Count);
    }

    public static function CommentAddByRowColSetsDefaults()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $comment = $sheet->Comments->Add(2, 3);
        AssertEx::Equal(1, $sheet->Comments->Count);
        AssertEx::Equal(2, $comment->Row);
        AssertEx::Equal(3, $comment->Column);
        AssertEx::Equal('', $comment->Note);
        AssertEx::Equal('', $comment->Author);
        AssertEx::False($comment->IsVisible);
        AssertEx::Equal(129, $comment->Width);
        AssertEx::Equal(75, $comment->Height);
    }

    public static function CommentAddByCellNameParsesReference()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $comment = $sheet->Comments->Add('B3');
        AssertEx::Equal(2, $comment->Row);
        AssertEx::Equal(1, $comment->Column);
        AssertEx::Equal(1, $sheet->Comments->Count);

        $byName = $sheet->Comments['B3'];
        AssertEx::NotNull($byName, 'Should find comment by cell name B3');
        AssertEx::Equal(2, $byName->Row);
    }

    public static function CommentAddDuplicateCellThrows()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Comments->Add(0, 0);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Comments->Add(0, 0);
        });
    }

    public static function CommentAddNegativeRowThrows()
    {
        $workbook = new Workbook();
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->Worksheets[0]->Comments->Add(-1, 0);
        });
    }

    public static function CommentAddNegativeColumnThrows()
    {
        $workbook = new Workbook();
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->Worksheets[0]->Comments->Add(0, -1);
        });
    }

    public static function CommentAddInvalidCellNameThrows()
    {
        $workbook = new Workbook();
        AssertEx::Throws(CellsException::class, function () use ($workbook) {
            $workbook->Worksheets[0]->Comments->Add('not_a_cell');
        });
    }

    public static function CommentCountIncrementsAfterAdd()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];

        $sheet->Comments->Add(0, 0);
        AssertEx::Equal(1, $sheet->Comments->Count);
        $sheet->Comments->Add(1, 0);
        AssertEx::Equal(2, $sheet->Comments->Count);
        $sheet->Comments->Add(0, 1);
        AssertEx::Equal(3, $sheet->Comments->Count);
    }

    public static function CommentRemoveAtIndexDecrementsCount()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Comments->Add(0, 0);
        $sheet->Comments->Add(1, 0);
        AssertEx::Equal(2, $sheet->Comments->Count);

        $sheet->Comments->RemoveAt(0);
        AssertEx::Equal(1, $sheet->Comments->Count);
        AssertEx::Equal(1, $sheet->Comments[0]->Row);

        $sheet->Comments->RemoveAt(0);
        AssertEx::Equal(0, $sheet->Comments->Count);
    }

    public static function CommentRemoveAtIndexOutOfRangeThrows()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Comments->RemoveAt(0);
        });
        $sheet->Comments->Add(0, 0);
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Comments->RemoveAt(1);
        });
        AssertEx::Throws(CellsException::class, function () use ($sheet) {
            $sheet->Comments->RemoveAt(-1);
        });
    }

    public static function CommentRemoveAtCellNameNoopWhenNotFound()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Comments->Add(0, 0);
        $sheet->Comments->RemoveAt('Z99');
        AssertEx::Equal(1, $sheet->Comments->Count);
    }

    public static function CommentIndexerByCellNameReturnsNullWhenNotFound()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Comments->Add(0, 0);
        $found = $sheet->Comments['A1'];
        AssertEx::NotNull($found, 'A1 should return the comment');
        $missing = $sheet->Comments['B2'];
        AssertEx::Null($missing, 'B2 has no comment; should return null');
    }

    public static function CommentNoteAndAuthorMutateViaProperty()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $comment = $sheet->Comments->Add(0, 0);

        $comment->Note = 'Review this';
        $comment->Author = 'Alice';
        $comment->IsVisible = true;

        AssertEx::Equal('Review this', $sheet->Comments[0]->Note);
        AssertEx::Equal('Alice', $sheet->Comments[0]->Author);
        AssertEx::True($sheet->Comments[0]->IsVisible);
    }

    public static function CommentIsVisibleDefaultsFalse()
    {
        $workbook = new Workbook();
        $comment = $workbook->Worksheets[0]->Comments->Add(0, 0);
        AssertEx::False($comment->IsVisible);
    }

    public static function CommentWidthHeightDefaultsAre129And75()
    {
        $workbook = new Workbook();
        $comment = $workbook->Worksheets[0]->Comments->Add(0, 0);
        AssertEx::Equal(129, $comment->Width);
        AssertEx::Equal(75, $comment->Height);
    }

    public static function CommentWidthZeroThrows()
    {
        $workbook = new Workbook();
        $comment = $workbook->Worksheets[0]->Comments->Add(0, 0);
        AssertEx::Throws(CellsException::class, function () use ($comment) {
            $comment->Width = 0;
        });
        AssertEx::Throws(CellsException::class, function () use ($comment) {
            $comment->Width = -5;
        });
    }

    public static function CommentHeightZeroThrows()
    {
        $workbook = new Workbook();
        $comment = $workbook->Worksheets[0]->Comments->Add(0, 0);
        AssertEx::Throws(CellsException::class, function () use ($comment) {
            $comment->Height = 0;
        });
        AssertEx::Throws(CellsException::class, function () use ($comment) {
            $comment->Height = -1;
        });
    }
}

Program::Main();
