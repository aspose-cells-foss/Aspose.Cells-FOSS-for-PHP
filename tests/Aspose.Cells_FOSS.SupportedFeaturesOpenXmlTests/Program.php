<?php

namespace Aspose\Cells_FOSS\SupportedFeaturesOpenXmlTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';
require_once __DIR__ . '/../Shared/ListObjectScenarioFactory.php';
require_once __DIR__ . '/../Shared/PictureScenarioFactory.php';
require_once __DIR__ . '/../Shared/ShapeScenarioFactory.php';
require_once __DIR__ . '/../Shared/WorkbookMetadataScenarioFactory.php';
require_once __DIR__ . '/ComparisonValueHelpers.php';
require_once __DIR__ . '/OpenXmlPackage.php';
require_once __DIR__ . '/SnapshotModels.php';
require_once __DIR__ . '/OpenXmlComparisonSupport.php';

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\PageOrientationType;
use Aspose\Cells_FOSS\PaperSizeType;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\ListObjectScenarioFactory;
use Aspose\Cells_FOSS\Testing\PictureScenarioFactory;
use Aspose\Cells_FOSS\Testing\ShapeScenarioFactory;
use Aspose\Cells_FOSS\Testing\StyleScenarioFactory;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\WorkbookMetadataScenarioFactory;
use Aspose\Cells_FOSS\Testing\ZipPackageHelper;
use Aspose\Cells_FOSS\ValidationAlertType;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\VisibilityType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    public static function Main()
    {
        return TestRunner::Run(
            'SupportedFeatures.OpenXmlTests',
            array(
                new TestCase('supported_features_workbook_matches_openxml_snapshot', array(Program::class, 'SupportedFeaturesWorkbookMatchesOpenXmlSnapshot')),
                new TestCase('workbook_metadata_matches_openxml_snapshot', array(Program::class, 'WorkbookMetadataMatchesOpenXmlSnapshot')),
                new TestCase('table_workbook_matches_openxml_snapshot', array(TableTestMethods::class, 'TableWorkbookMatchesOpenXmlSnapshot')),
                new TestCase('table_workbook_roundtrip', array(TableTestMethods::class, 'TableWorkbookRoundtrip')),
                new TestCase('picture_workbook_matches_openxml_snapshot', array(PictureTestMethods::class, 'PictureWorkbookMatchesOpenXmlSnapshot')),
                new TestCase('picture_workbook_roundtrip', array(PictureTestMethods::class, 'PictureWorkbookRoundtrip')),
                new TestCase('shape_workbook_matches_openxml_snapshot', array(ShapeTestMethods::class, 'ShapeWorkbookMatchesOpenXmlSnapshot')),
                new TestCase('shape_workbook_roundtrip', array(ShapeTestMethods::class, 'ShapeWorkbookRoundtrip')),
            )
        );
    }

    public static function SupportedFeaturesWorkbookMatchesOpenXmlSnapshot()
    {
        $temp = new TemporaryDirectory('supported-features-openxml');
        $path = $temp->GetPath('supported-features.xlsx');

        $workbook = self::CreateSupportedFeaturesWorkbook();
        $saveOptions = new \Aspose\Cells_FOSS\SaveOptions();
        $saveOptions->UseSharedStrings = true;
        $workbook->Save($path, $saveOptions);

        $snapshot = self::ReadWorkbookSnapshot($path);

        AssertEx::True($snapshot->UsesDate1904);
        AssertEx::True($snapshot->HasSharedStrings);
        AssertEx::Equal(3, count($snapshot->Sheets));
        self::AssertSheet($snapshot->Sheets[0], 'All Features', 'Visible');
        self::AssertSheet($snapshot->Sheets[1], 'Target Sheet', 'Hidden');
        self::AssertSheet($snapshot->Sheets[2], 'VeryHidden Sheet', 'VeryHidden');

        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'A1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'B1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'C1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'D1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'E1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'F1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'G1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'H1');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'I2');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'All Features', 'J3');
        self::AssertExpectedCell($workbook, $snapshot->Cells, 'Target Sheet', 'C3');

        self::AssertExpectedStyle($workbook, $snapshot->Styles, 'All Features', 'A1');
        self::AssertExpectedStyle($workbook, $snapshot->Styles, 'All Features', 'C4');
        self::AssertExpectedStyle($workbook, $snapshot->Styles, 'All Features', 'H1');
        self::AssertExpectedStyle($workbook, $snapshot->Styles, 'All Features', 'I2');
        self::AssertExpectedStyle($workbook, $snapshot->Styles, 'All Features', 'J3');

        $worksheet = $snapshot->Worksheets['All Features'];
        AssertEx::Equal('A1:J4', $worksheet->Dimension);
        self::AssertRow($worksheet, 1, 22.5, false);
        self::AssertRow($worksheet, 3, null, true);
        self::AssertColumn($worksheet, 0, 0, 18.25, false);
        self::AssertColumn($worksheet, 2, 2, null, true);
        AssertEx::Equal(1, count($worksheet->MergeReferences));
        AssertEx::Equal('A3:B4', $worksheet->MergeReferences[0]);

        AssertEx::Equal(3, count($worksheet->Hyperlinks));
        self::AssertHyperlink($worksheet->Hyperlinks[0], 'H1', 'https://example.com/docs?q=1', '', 'External docs', 'Open External Documentation');
        self::AssertHyperlink($worksheet->Hyperlinks[1], 'I2', '', "'Target Sheet'!C3", 'Jump to target', 'Navigate to Target Sheet');
        self::AssertHyperlink($worksheet->Hyperlinks[2], 'J3:K4', 'mailto:test@example.com', '', 'Send mail', 'Contact Us by Email');

        AssertEx::Equal(3, count($worksheet->Validations));
        self::AssertValidation(self::GetValidationSnapshot($worksheet->Validations, 'A1:A3'), 'A1:A3', 'list', '', 'stop', true, false, true, true, 'Status', 'Pick a status', 'Invalid', 'Choose from the list', '"Open,Closed"', '');
        self::AssertValidation(self::GetValidationSnapshot($worksheet->Validations, 'B2:C3 E2:E3'), 'B2:C3 E2:E3', 'decimal', 'between', 'stop', false, true, false, true, '', '', 'Range', 'Enter 1.5-9.5', '1.5', '9.5');
        self::AssertValidation(self::GetValidationSnapshot($worksheet->Validations, 'G1'), 'G1', 'custom', '', 'warning', false, false, true, false, 'Code', 'Up to 5 chars', '', '', 'LEN(G1)<=5', '');

        AssertEx::Equal(2, count($worksheet->ConditionalFormattings));
        self::AssertConditionalFormatting(self::GetConditionalFormattingSnapshot($worksheet->ConditionalFormattings, 'A1:A5'), 'A1:A5', 2);
        self::AssertConditionalFormatting(self::GetConditionalFormattingSnapshot($worksheet->ConditionalFormattings, 'C1:C3'), 'C1:C3', 1);

        $pageSetup = $worksheet->PageSetup;
        AssertEx::Equal('$A$1:$J$10', $pageSetup->PrintArea);
        AssertEx::Equal('$1:$2', $pageSetup->PrintTitleRows);
        AssertEx::Equal('$A:$B', $pageSetup->PrintTitleColumns);
        AssertEx::Equal(0.25, $pageSetup->LeftMargin === null ? 0.0 : $pageSetup->LeftMargin);
        AssertEx::Equal(0.4, $pageSetup->RightMargin === null ? 0.0 : $pageSetup->RightMargin);
        AssertEx::Equal(0.5, $pageSetup->TopMargin === null ? 0.0 : $pageSetup->TopMargin);
        AssertEx::Equal(0.6, $pageSetup->BottomMargin === null ? 0.0 : $pageSetup->BottomMargin);
        AssertEx::Equal(0.2, $pageSetup->HeaderMargin === null ? 0.0 : $pageSetup->HeaderMargin);
        AssertEx::Equal(0.22, $pageSetup->FooterMargin === null ? 0.0 : $pageSetup->FooterMargin);
        AssertEx::Equal(9, $pageSetup->PaperSize === null ? 0 : $pageSetup->PaperSize);
        AssertEx::Equal('landscape', $pageSetup->Orientation);
        AssertEx::Equal(3, $pageSetup->FirstPageNumber === null ? 0 : $pageSetup->FirstPageNumber);
        AssertEx::Equal(95, $pageSetup->Scale === null ? 0 : $pageSetup->Scale);
        AssertEx::Equal(1, $pageSetup->FitToWidth === null ? 0 : $pageSetup->FitToWidth);
        AssertEx::Equal(2, $pageSetup->FitToHeight === null ? 0 : $pageSetup->FitToHeight);
        AssertEx::True($pageSetup->PrintGridLines);
        AssertEx::True($pageSetup->PrintHeadings);
        AssertEx::True($pageSetup->HorizontalCentered);
        AssertEx::True($pageSetup->VerticalCentered);
        AssertEx::Equal('&LLeft Header&CCenter Header&RRight Header', $pageSetup->OddHeader);
        AssertEx::Equal('&LLeft Footer&CCenter Footer&RRight Footer', $pageSetup->OddFooter);
        AssertEx::Equal(2, count($pageSetup->RowBreaks));
        AssertEx::Equal(4, $pageSetup->RowBreaks[0]);
        AssertEx::Equal(7, $pageSetup->RowBreaks[1]);
        AssertEx::Equal(1, count($pageSetup->ColumnBreaks));
        AssertEx::Equal(2, $pageSetup->ColumnBreaks[0]);

        $definedNames = self::ReadDefinedNames($path);
        AssertEx::Equal(2, count($definedNames));
        self::AssertDefinedName($definedNames[0], 'GlobalAmount', "'All Features'!\$A\$1:\$A\$3", null, true, 'Primary selection');
        self::AssertDefinedName($definedNames[1], 'LocalTarget', "'Target Sheet'!\$C\$3", 1, false, '');
    }

    public static function WorkbookMetadataMatchesOpenXmlSnapshot()
    {
        $temp = new TemporaryDirectory('workbook-metadata-openxml');
        $path = $temp->GetPath('workbook-metadata.xlsx');

        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($path);

        $package = new OpenXmlPackage($path);
        $workbookXml = $package->ReadText('xl/workbook.xml');
        $workbookDocument = XDocument::Load($workbookXml);
        $workbookRoot = $workbookDocument->Root;

        $workbookProperties = $workbookRoot->Element(self::MainNs . 'workbookPr');
        AssertEx::Equal('WorkbookCode', $workbookProperties === null ? null : self::AttributeText($workbookProperties, 'codeName'));
        AssertEx::NotNull($workbookProperties === null ? null : self::AttributeText($workbookProperties, 'showObjects'));
        AssertEx::NotNull($workbookRoot->Element(self::MainNs . 'workbookProtection'));

        $bookViewsElement = $workbookRoot->Element(self::MainNs . 'bookViews');
        $firstBookView = null;
        if ($bookViewsElement !== null) {
            $views = $bookViewsElement->Elements(self::MainNs . 'workbookView');
            if (count($views) > 0) {
                $firstBookView = $views[0];
            }
        }

        AssertEx::Equal(1, $firstBookView === null ? 0 : (int) self::AttributeText($firstBookView, 'activeTab'));
        AssertEx::Equal(1, $firstBookView === null ? 0 : (int) self::AttributeText($firstBookView, 'firstSheet'));

        $calculationProperties = $workbookRoot->Element(self::MainNs . 'calcPr');
        AssertEx::NotNull($calculationProperties === null ? null : self::AttributeText($calculationProperties, 'calcMode'));
        AssertEx::NotNull($calculationProperties === null ? null : self::AttributeText($calculationProperties, 'refMode'));

        $coreXml = ZipPackageHelper::ReadEntryText($path, 'docProps/core.xml');
        $appXml = ZipPackageHelper::ReadEntryText($path, 'docProps/app.xml');
        AssertEx::Contains('showObjects="placeholders"', $workbookXml);
        AssertEx::Contains('calcMode="manual"', $workbookXml);
        AssertEx::Contains('refMode="R1C1"', $workbookXml);
        AssertEx::Contains('Quarterly Summary', $coreXml);
        AssertEx::Contains('Automation', $coreXml);
        AssertEx::Contains('Aspose.Cells_FOSS Tests', $appXml);
        AssertEx::Contains('https://example.com/base/', $appXml);

        $loaded = new Workbook($path);
        WorkbookMetadataScenarioFactory::AssertWorkbookMetadata($loaded);
    }

    private static function CreateSupportedFeaturesWorkbook()
    {
        $workbook = new Workbook();
        $workbook->Settings->Date1904 = true;

        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'All Features';

        $targetIndex = $workbook->Worksheets->Add('Target Sheet');
        $targetSheet = $workbook->Worksheets[$targetIndex];
        $targetSheet->VisibilityType = VisibilityType::Hidden;
        $targetSheet->Cells[[2, 2]]->PutValue('Target');

        $veryHiddenIndex = $workbook->Worksheets->Add('VeryHidden Sheet');
        $workbook->Worksheets[$veryHiddenIndex]->VisibilityType = VisibilityType::VeryHidden;

        $sheet->Cells['A1']->PutValue('Hello');
        $sheet->Cells['B1']->PutValue(123);
        $sheet->Cells['C1']->PutValue(true);
        $sheet->Cells['D1']->PutValue(12.5);
        $sheet->Cells['E1']->PutValue(6.02214076E+23);
        $sheet->Cells['F1']->PutValue(new \DateTime('2024-05-06 07:08:09', new \DateTimeZone('UTC')));
        $sheet->Cells['G1']->PutValue(20);
        $sheet->Cells['G1']->Formula = '=B1*2';
        $sheet->Cells['A3']->PutValue('Merged');
        $sheet->Cells['C4']->PutValue(42.1234);

        $primaryStyle = $sheet->Cells['A1']->GetStyle();
        StyleScenarioFactory::ApplyPrimaryStyle($primaryStyle);
        $sheet->Cells['A1']->SetStyle($primaryStyle);

        $customStyle = $sheet->Cells['C4']->GetStyle();
        StyleScenarioFactory::ApplyCustomNumberStyle($customStyle);
        $sheet->Cells['C4']->SetStyle($customStyle);

        $sheet->Cells->Rows[1]->Height = 22.5;
        $sheet->Cells->Rows[3]->IsHidden = true;
        $sheet->Cells->Columns[0]->Width = 18.25;
        $sheet->Cells->Columns[2]->IsHidden = true;
        $sheet->Cells->Merge(2, 0, 2, 2);

        $hyperlinkColor = Color::FromArgb(255, 5, 99, 193);

        $sheet->Cells['H1']->PutValue('Open External Documentation');
        $external = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('H1', 1, 1, 'https://example.com/docs?q=1')];
        $external->TextToDisplay = 'Open External Documentation';
        $external->ScreenTip = 'External docs';
        $h1Style = $sheet->Cells['H1']->GetStyle();
        $h1Style->Font->Color = $hyperlinkColor;
        $h1Style->Font->Underline = FontUnderlineType::Single;
        $sheet->Cells['H1']->SetStyle($h1Style);

        $sheet->Cells['I2']->PutValue('Navigate to Target Sheet');
        $internalLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('I2', 1, 1, "'Target Sheet'!C3")];
        $internalLink->TextToDisplay = 'Navigate to Target Sheet';
        $internalLink->ScreenTip = 'Jump to target';
        $i2Style = $sheet->Cells['I2']->GetStyle();
        $i2Style->Font->Color = $hyperlinkColor;
        $i2Style->Font->Underline = FontUnderlineType::Single;
        $sheet->Cells['I2']->SetStyle($i2Style);

        $sheet->Cells['J3']->PutValue('Contact Us by Email');
        $rangeLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('J3', 2, 2, 'mailto:test@example.com')];
        $rangeLink->TextToDisplay = 'Contact Us by Email';
        $rangeLink->ScreenTip = 'Send mail';
        $j3Style = $sheet->Cells['J3']->GetStyle();
        $j3Style->Font->Color = $hyperlinkColor;
        $j3Style->Font->Underline = FontUnderlineType::Single;
        $sheet->Cells['J3']->SetStyle($j3Style);

        $listValidation = $sheet->Validations[$sheet->Validations->Add(\Aspose\Cells_FOSS\CellArea::CreateCellArea('A1', 'A3'))];
        $listValidation->Type = ValidationType::List;
        $listValidation->Formula1 = '"Open,Closed"';
        $listValidation->IgnoreBlank = true;
        $listValidation->ShowInput = true;
        $listValidation->InputTitle = 'Status';
        $listValidation->InputMessage = 'Pick a status';
        $listValidation->ShowError = true;
        $listValidation->ErrorTitle = 'Invalid';
        $listValidation->ErrorMessage = 'Choose from the list';

        $decimalValidation = $sheet->Validations[$sheet->Validations->Add(\Aspose\Cells_FOSS\CellArea::CreateCellArea('B2', 'C3'))];
        $decimalValidation->Type = ValidationType::Decimal;
        $decimalValidation->Operator = OperatorType::Between;
        $decimalValidation->Formula1 = '1.5';
        $decimalValidation->Formula2 = '9.5';
        $decimalValidation->InCellDropDown = false;
        $decimalValidation->ShowError = true;
        $decimalValidation->ErrorTitle = 'Range';
        $decimalValidation->ErrorMessage = 'Enter 1.5-9.5';
        $decimalValidation->AddArea(\Aspose\Cells_FOSS\CellArea::CreateCellArea('E2', 'E3'));

        $customValidation = $sheet->Validations[$sheet->Validations->Add(\Aspose\Cells_FOSS\CellArea::CreateCellArea('G1', 'G1'))];
        $customValidation->Type = ValidationType::Custom;
        $customValidation->AlertStyle = ValidationAlertType::Warning;
        $customValidation->Formula1 = 'LEN(G1)<=5';
        $customValidation->ShowInput = true;
        $customValidation->InputTitle = 'Code';
        $customValidation->InputMessage = 'Up to 5 chars';

        $primaryConditionalFormatting = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $primaryConditionalFormatting->AddArea(\Aspose\Cells_FOSS\CellArea::CreateCellArea('A1', 'A5'));
        $betweenCondition = $primaryConditionalFormatting[$primaryConditionalFormatting->AddCondition(FormatConditionType::CellValue, OperatorType::Between, '5', '15')];
        $betweenCondition->StopIfTrue = true;
        $betweenStyle = $betweenCondition->Style;
        $betweenStyle->Pattern = \Aspose\Cells_FOSS\FillPattern::Solid;
        $betweenStyle->ForegroundColor = Color::FromArgb(255, 255, 199, 206);
        $betweenStyle->Font->IsBold = true;
        $betweenStyle->Font->Color = Color::FromArgb(255, 156, 0, 6);
        $betweenCondition->Style = $betweenStyle;
        $expressionCondition = $primaryConditionalFormatting[$primaryConditionalFormatting->AddCondition(FormatConditionType::Expression, OperatorType::None, 'MOD(A1,2)=0', '')];
        $expressionStyle = $expressionCondition->Style;
        $expressionStyle->Font->IsItalic = true;
        $expressionStyle->Font->Color = Color::FromArgb(255, 0, 0, 255);
        $expressionCondition->Style = $expressionStyle;

        $secondaryConditionalFormatting = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $secondaryConditionalFormatting->AddArea(\Aspose\Cells_FOSS\CellArea::CreateCellArea('C1', 'C3'));
        $greaterThanCondition = $secondaryConditionalFormatting[$secondaryConditionalFormatting->AddCondition(FormatConditionType::CellValue, OperatorType::GreaterThan, '10', '')];
        $greaterThanStyle = $greaterThanCondition->Style;
        $greaterThanStyle->Pattern = \Aspose\Cells_FOSS\FillPattern::Solid;
        $greaterThanStyle->ForegroundColor = Color::FromArgb(255, 198, 239, 206);
        $greaterThanStyle->Font->Color = Color::FromArgb(255, 0, 97, 0);
        $greaterThanCondition->Style = $greaterThanStyle;

        $pageSetup = $sheet->PageSetup;
        $pageSetup->LeftMarginInch = 0.25;
        $pageSetup->RightMarginInch = 0.4;
        $pageSetup->TopMarginInch = 0.5;
        $pageSetup->BottomMarginInch = 0.6;
        $pageSetup->HeaderMarginInch = 0.2;
        $pageSetup->FooterMarginInch = 0.22;
        $pageSetup->Orientation = PageOrientationType::Landscape;
        $pageSetup->PaperSize = PaperSizeType::PaperA4;
        $pageSetup->FirstPageNumber = 3;
        $pageSetup->Scale = 95;
        $pageSetup->FitToPagesWide = 1;
        $pageSetup->FitToPagesTall = 2;
        $pageSetup->PrintArea = '$A$1:$J$10';
        $pageSetup->PrintTitleRows = '$1:$2';
        $pageSetup->PrintTitleColumns = '$A:$B';
        $pageSetup->LeftHeader = 'Left Header';
        $pageSetup->CenterHeader = 'Center Header';
        $pageSetup->RightHeader = 'Right Header';
        $pageSetup->LeftFooter = 'Left Footer';
        $pageSetup->CenterFooter = 'Center Footer';
        $pageSetup->RightFooter = 'Right Footer';
        $pageSetup->PrintGridlines = true;
        $pageSetup->PrintHeadings = true;
        $pageSetup->CenterHorizontally = true;
        $pageSetup->CenterVertically = true;
        $pageSetup->AddHorizontalPageBreak(4);
        $pageSetup->AddHorizontalPageBreak(7);
        $pageSetup->AddVerticalPageBreak(2);

        $globalDefinedName = $workbook->DefinedNames[$workbook->DefinedNames->Add('GlobalAmount', "'All Features'!\$A\$1:\$A\$3")];
        $globalDefinedName->Hidden = true;
        $globalDefinedName->Comment = 'Primary selection';
        $workbook->DefinedNames->Add('LocalTarget', "'Target Sheet'!\$C\$3", 1);

        return $workbook;
    }

    private static function AssertExpectedCell($workbook, $actualCells, $sheetName, $cellReference)
    {
        $sheet = $workbook->Worksheets[$sheetName];
        $cell = $sheet->Cells[$cellReference];
        $expected = new CellSnapshot(
            $sheetName,
            $cellReference,
            self::NormalizeCellType($cell->Type),
            ComparisonValueHelpers::NormalizeValue($cell->Value),
            ComparisonValueHelpers::NormalizeFormulaText($cell->Formula)
        );
        $key = ComparisonValueHelpers::BuildCellKey($sheetName, $cellReference);
        AssertEx::True(array_key_exists($key, $actualCells), 'Expected Open XML cell snapshot for ' . $key . '.');
        $actual = $actualCells[$key];

        AssertEx::Equal($expected->SheetName, $actual->SheetName);
        AssertEx::Equal($expected->CellReference, $actual->CellReference);
        AssertEx::Equal($expected->CellType, $actual->CellType);
        AssertEx::Equal($expected->Formula, $actual->Formula);

        if ($cell->Value instanceof \DateTime) {
            $expectedUtc = clone $cell->Value;
            $expectedUtc->setTimezone(new \DateTimeZone('UTC'));
            AssertEx::Equal($expectedUtc->format('Y-m-d\TH:i:s'), substr($actual->Value, 0, 19));
            return;
        }

        AssertEx::Equal($expected->Value, $actual->Value);
    }

    private static function AssertExpectedStyle($workbook, $actualStyles, $sheetName, $cellReference)
    {
        $expected = OpenXmlComparisonSupport::CreateLibraryStyleSnapshot(
            $sheetName,
            $cellReference,
            $workbook->Worksheets[$sheetName]->Cells[$cellReference]->GetStyle()
        );
        $key = ComparisonValueHelpers::BuildCellKey($sheetName, $cellReference);
        AssertEx::True(array_key_exists($key, $actualStyles), 'Expected Open XML style snapshot for ' . $key . '.');
        $actual = $actualStyles[$key];

        AssertEx::Equal($expected->SheetName, $actual->SheetName);
        AssertEx::Equal($expected->CellReference, $actual->CellReference);
        AssertEx::Equal($expected->FontName, $actual->FontName);
        AssertEx::Equal($expected->FontSize, $actual->FontSize);
        AssertEx::Equal($expected->FontBold, $actual->FontBold);
        AssertEx::Equal($expected->FontItalic, $actual->FontItalic);
        AssertEx::Equal($expected->FontUnderline, $actual->FontUnderline);
        AssertEx::Equal($expected->FontStrikeThrough, $actual->FontStrikeThrough);
        AssertEx::Equal($expected->FontColor, $actual->FontColor);
        AssertEx::Equal($expected->FillPattern, $actual->FillPattern);
        AssertEx::Equal($expected->FillForegroundColor, $actual->FillForegroundColor);
        AssertEx::Equal($expected->FillBackgroundColor, $actual->FillBackgroundColor);
        AssertEx::Equal($expected->LeftBorderStyle, $actual->LeftBorderStyle);
        AssertEx::Equal($expected->LeftBorderColor, $actual->LeftBorderColor);
        AssertEx::Equal($expected->RightBorderStyle, $actual->RightBorderStyle);
        AssertEx::Equal($expected->RightBorderColor, $actual->RightBorderColor);
        AssertEx::Equal($expected->TopBorderStyle, $actual->TopBorderStyle);
        AssertEx::Equal($expected->TopBorderColor, $actual->TopBorderColor);
        AssertEx::Equal($expected->BottomBorderStyle, $actual->BottomBorderStyle);
        AssertEx::Equal($expected->BottomBorderColor, $actual->BottomBorderColor);
        AssertEx::Equal($expected->DiagonalBorderStyle, $actual->DiagonalBorderStyle);
        AssertEx::Equal($expected->DiagonalBorderColor, $actual->DiagonalBorderColor);
        AssertEx::Equal($expected->DiagonalUp, $actual->DiagonalUp);
        AssertEx::Equal($expected->DiagonalDown, $actual->DiagonalDown);
        AssertEx::Equal($expected->HorizontalAlignment, $actual->HorizontalAlignment);
        AssertEx::Equal($expected->VerticalAlignment, $actual->VerticalAlignment);
        AssertEx::Equal($expected->WrapText, $actual->WrapText);
        AssertEx::Equal($expected->IndentLevel, $actual->IndentLevel);
        AssertEx::Equal($expected->TextRotation, $actual->TextRotation);
        AssertEx::Equal($expected->ShrinkToFit, $actual->ShrinkToFit);
        AssertEx::Equal($expected->ReadingOrder, $actual->ReadingOrder);
        AssertEx::Equal($expected->RelativeIndent, $actual->RelativeIndent);
        AssertEx::Equal($expected->IsLocked, $actual->IsLocked);
        AssertEx::Equal($expected->IsHidden, $actual->IsHidden);

        if ($expected->NumberFormatCode !== null && $expected->NumberFormatCode !== '') {
            AssertEx::Equal($expected->NumberFormatCode, $actual->NumberFormatCode);
        } else {
            AssertEx::Equal($expected->NumberFormatId, $actual->NumberFormatId);
        }
    }

    private static function NormalizeCellType($type)
    {
        switch ($type) {
            case \Aspose\Cells_FOSS\CellValueType::IsString:
                return 'String';
            case \Aspose\Cells_FOSS\CellValueType::IsNumeric:
                return 'Number';
            case \Aspose\Cells_FOSS\CellValueType::IsBool:
                return 'Boolean';
            case \Aspose\Cells_FOSS\CellValueType::IsDateTime:
                return 'DateTime';
            case \Aspose\Cells_FOSS\CellValueType::IsUnknown:
                return 'Formula';
            default:
                return 'Blank';
        }
    }

    private static function ReadWorkbookSnapshot($workbookPath)
    {
        $cells = OpenXmlComparisonSupport::ReadCellDataWithOpenXmlSdk($workbookPath);
        $styles = OpenXmlComparisonSupport::ReadCellStylesWithOpenXmlSdk($workbookPath);

        $package = new OpenXmlPackage($workbookPath);
        $workbookDocument = $package->ReadXml('xl/workbook.xml');
        $workbookRoot = $workbookDocument->Root;
        if ($workbookRoot === null) {
            throw new \RuntimeException('Workbook XML is empty.');
        }

        $relationships = $package->ReadRelationships('xl/workbook.xml');
        $sheets = array();
        $worksheetSnapshots = array();

        $sheetsElement = $workbookRoot->Element(self::MainNs . 'sheets');
        if ($sheetsElement !== null) {
            foreach ($sheetsElement->Elements(self::MainNs . 'sheet') as $sheetElement) {
                $relationshipId = self::AttributeText($sheetElement, self::RelationshipNs . 'id');
                if ($relationshipId === null || trim($relationshipId) === '' || !array_key_exists($relationshipId, $relationships)) {
                    continue;
                }

                $target = $relationships[$relationshipId]['ResolvedTarget'];
                if ($target === null || $target === '') {
                    continue;
                }

                $part = ltrim($target, '/');
                if (!$package->Exists($part)) {
                    continue;
                }

                $sheetName = self::AttributeText($sheetElement, 'name');
                if ($sheetName === null) {
                    $sheetName = '';
                }

                $sheets[] = new SheetSnapshot($sheetName, self::ResolveVisibility(self::AttributeText($sheetElement, 'state')));
                $worksheetSnapshots[$sheetName] = self::ReadWorksheetSnapshot($package, $workbookRoot, $sheetElement, $part);
            }
        }

        $workbookProperties = $workbookRoot->Element(self::MainNs . 'workbookPr');
        $usesDate1904 = ComparisonValueHelpers::ParseBoolAttribute($workbookProperties === null ? null : $workbookProperties->Attribute('date1904'));
        $hasSharedStrings = self::ResolvePart($relationships, '/sharedStrings') !== null;

        $snapshot = new WorkbookSnapshot();
        $snapshot->UsesDate1904 = $usesDate1904;
        $snapshot->HasSharedStrings = $hasSharedStrings;
        $snapshot->Sheets = $sheets;
        $snapshot->Cells = $cells;
        $snapshot->Styles = $styles;
        $snapshot->Worksheets = $worksheetSnapshots;
        return $snapshot;
    }

    private static function ReadWorksheetSnapshot($package, $workbookRoot, $sheetElement, $worksheetPart)
    {
        $worksheetDocument = $package->ReadXml($worksheetPart);
        $worksheetRoot = $worksheetDocument->Root;
        if ($worksheetRoot === null) {
            throw new \RuntimeException('Worksheet XML is empty.');
        }

        $sheetName = self::AttributeText($sheetElement, 'name');
        if ($sheetName === null) {
            $sheetName = '';
        }

        $sheetIdText = self::AttributeText($sheetElement, 'sheetId');
        $sheetIndex = ($sheetIdText === null ? 1 : (int) $sheetIdText) - 1;

        $rows = array();
        $sheetData = $worksheetRoot->Element(self::MainNs . 'sheetData');
        if ($sheetData !== null) {
            foreach ($sheetData->Elements(self::MainNs . 'row') as $row) {
                $rowIndex = ComparisonValueHelpers::ParseIntValue(self::AttributeText($row, 'r'));
                if ($rowIndex === null) {
                    continue;
                }

                $rows[$rowIndex - 1] = new RowSnapshot(
                    ComparisonValueHelpers::ParseDoubleValueOrNull(self::AttributeText($row, 'ht')),
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($row, 'hidden'))
                );
            }
        }

        $columns = array();
        $cols = $worksheetRoot->Element(self::MainNs . 'cols');
        if ($cols !== null) {
            foreach ($cols->Elements(self::MainNs . 'col') as $column) {
                $minColumnIndex = ComparisonValueHelpers::ParseIntValue(self::AttributeText($column, 'min'));
                $maxColumnIndex = ComparisonValueHelpers::ParseIntValue(self::AttributeText($column, 'max'));
                if ($minColumnIndex === null || $maxColumnIndex === null) {
                    continue;
                }

                $columns[] = new ColumnSnapshot(
                    $minColumnIndex - 1,
                    $maxColumnIndex - 1,
                    ComparisonValueHelpers::ParseDoubleValueOrNull(self::AttributeText($column, 'width')),
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($column, 'hidden'))
                );
            }
        }

        $merges = array();
        $mergeCells = $worksheetRoot->Element(self::MainNs . 'mergeCells');
        if ($mergeCells !== null) {
            foreach ($mergeCells->Elements(self::MainNs . 'mergeCell') as $mergeCell) {
                $reference = strtoupper(self::AttributeText($mergeCell, 'ref'));
                if ($reference !== '') {
                    $merges[] = $reference;
                }
            }
        }

        $hyperlinkTargets = array();
        foreach ($package->ReadRelationships($worksheetPart) as $relationshipId => $relationship) {
            if (strpos($relationship['Type'], '/hyperlink') === false) {
                continue;
            }

            $hyperlinkTargets[$relationshipId] = $relationship['IsExternal'] ? $relationship['Target'] : $relationship['ResolvedTarget'];
        }

        $hyperlinks = array();
        $hyperlinksElement = $worksheetRoot->Element(self::MainNs . 'hyperlinks');
        if ($hyperlinksElement !== null) {
            foreach ($hyperlinksElement->Elements(self::MainNs . 'hyperlink') as $hyperlink) {
                $relationshipId = self::AttributeText($hyperlink, self::RelationshipNs . 'id');
                $hyperlinkAddress = '';
                if ($relationshipId !== null && array_key_exists($relationshipId, $hyperlinkTargets)) {
                    $hyperlinkAddress = $hyperlinkTargets[$relationshipId];
                }

                $hyperlinks[] = new HyperlinkSnapshot(
                    strtoupper(self::AttributeText($hyperlink, 'ref')),
                    $hyperlinkAddress,
                    self::AttributeTextOrEmpty($hyperlink, 'location'),
                    self::AttributeTextOrEmpty($hyperlink, 'tooltip'),
                    self::AttributeTextOrEmpty($hyperlink, 'display')
                );
            }
        }

        $validations = array();
        $dataValidations = $worksheetRoot->Element(self::MainNs . 'dataValidations');
        if ($dataValidations !== null) {
            foreach ($dataValidations->Elements(self::MainNs . 'dataValidation') as $validation) {
                $errorStyle = self::AttributeText($validation, 'errorStyle');
                if ($errorStyle === null || $errorStyle === '') {
                    $errorStyle = 'stop';
                }

                $validations[] = new ValidationSnapshot(
                    strtoupper(self::AttributeText($validation, 'sqref')),
                    self::AttributeTextOrEmpty($validation, 'type'),
                    self::AttributeTextOrEmpty($validation, 'operator'),
                    $errorStyle,
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($validation, 'allowBlank')),
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($validation, 'showDropDown')),
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($validation, 'showInputMessage')),
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($validation, 'showErrorMessage')),
                    self::AttributeTextOrEmpty($validation, 'promptTitle'),
                    self::AttributeTextOrEmpty($validation, 'prompt'),
                    self::AttributeTextOrEmpty($validation, 'errorTitle'),
                    self::AttributeTextOrEmpty($validation, 'error'),
                    self::ChildTextOrEmpty($validation, self::MainNs . 'formula1'),
                    self::ChildTextOrEmpty($validation, self::MainNs . 'formula2')
                );
            }
        }

        $conditionalFormattings = array();
        foreach ($worksheetRoot->Elements(self::MainNs . 'conditionalFormatting') as $formatting) {
            $rules = array();
            foreach ($formatting->Elements(self::MainNs . 'cfRule') as $rule) {
                $formulas = array();
                foreach ($rule->Elements(self::MainNs . 'formula') as $formula) {
                    $formulas[] = $formula->Value;
                }

                $rules[] = new ConditionalFormattingRuleSnapshot(
                    self::AttributeTextOrEmpty($rule, 'type'),
                    self::AttributeTextOrEmpty($rule, 'operator'),
                    self::AttributeTextOrEmpty($rule, 'priority'),
                    self::AttributeTextOrEmpty($rule, 'stopIfTrue'),
                    self::AttributeTextOrEmpty($rule, 'dxfId'),
                    $formulas
                );
            }

            $conditionalFormattings[] = new ConditionalFormattingSnapshot(
                strtoupper(self::AttributeText($formatting, 'sqref')),
                $rules
            );
        }

        $pageSetup = self::ReadPageSetupSnapshot($workbookRoot, $worksheetRoot, $sheetName, $sheetIndex);

        return new WorksheetSnapshot(
            $sheetName,
            self::ResolveVisibility(self::AttributeText($sheetElement, 'state')),
            strtoupper(self::ChildAttributeText($worksheetRoot, self::MainNs . 'dimension', 'ref')),
            $rows,
            $columns,
            $merges,
            $hyperlinks,
            $validations,
            $conditionalFormattings,
            $pageSetup
        );
    }

    private static function ReadPageSetupSnapshot($workbookRoot, $worksheetRoot, $sheetName, $sheetIndex)
    {
        $margins = $worksheetRoot->Element(self::MainNs . 'pageMargins');
        $pageSetup = $worksheetRoot->Element(self::MainNs . 'pageSetup');
        $printOptions = $worksheetRoot->Element(self::MainNs . 'printOptions');
        $headerFooter = $worksheetRoot->Element(self::MainNs . 'headerFooter');

        $rowBreaks = array();
        $rowBreaksElement = $worksheetRoot->Element(self::MainNs . 'rowBreaks');
        if ($rowBreaksElement !== null) {
            foreach ($rowBreaksElement->Elements(self::MainNs . 'brk') as $brk) {
                $rowBreaks[] = self::ParseUInt(self::AttributeText($brk, 'id'));
            }
        }

        $columnBreaks = array();
        $colBreaksElement = $worksheetRoot->Element(self::MainNs . 'colBreaks');
        if ($colBreaksElement !== null) {
            foreach ($colBreaksElement->Elements(self::MainNs . 'brk') as $brk) {
                $columnBreaks[] = self::ParseUInt(self::AttributeText($brk, 'id'));
            }
        }

        $printArea = null;
        $printTitleRows = null;
        $printTitleColumns = null;
        $definedNamesElement = $workbookRoot->Element(self::MainNs . 'definedNames');
        if ($definedNamesElement !== null) {
            foreach ($definedNamesElement->Elements(self::MainNs . 'definedName') as $definedName) {
                $localSheetId = self::AttributeText($definedName, 'localSheetId');
                if ($localSheetId !== (string) $sheetIndex) {
                    continue;
                }

                $name = self::AttributeTextOrEmpty($definedName, 'name');
                if ($name === '_xlnm.Print_Area') {
                    $printArea = self::UnqualifyDefinedNameValue($definedName->Value);
                } elseif ($name === '_xlnm.Print_Titles') {
                    foreach (explode(',', $definedName->Value) as $segment) {
                        $unqualified = self::UnqualifyDefinedNameValue($segment);
                        $withoutDollar = str_replace('$', '', $unqualified);
                        $hasDigit = preg_match('/[0-9]/', $unqualified) === 1;
                        $firstIsDigit = $withoutDollar !== '' && $withoutDollar[0] >= '0' && $withoutDollar[0] <= '9';
                        if (strpos($unqualified, '$') !== false && $hasDigit && strpos($unqualified, ':') !== false && $firstIsDigit) {
                            $printTitleRows = $unqualified;
                        } else {
                            $printTitleColumns = $unqualified;
                        }
                    }
                }
            }
        }

        return new PageSetupSnapshot(
            $printArea === null ? '' : $printArea,
            $printTitleRows === null ? '' : $printTitleRows,
            $printTitleColumns === null ? '' : $printTitleColumns,
            ComparisonValueHelpers::ParseDoubleValueOrNull($margins === null ? null : self::AttributeText($margins, 'left')),
            ComparisonValueHelpers::ParseDoubleValueOrNull($margins === null ? null : self::AttributeText($margins, 'right')),
            ComparisonValueHelpers::ParseDoubleValueOrNull($margins === null ? null : self::AttributeText($margins, 'top')),
            ComparisonValueHelpers::ParseDoubleValueOrNull($margins === null ? null : self::AttributeText($margins, 'bottom')),
            ComparisonValueHelpers::ParseDoubleValueOrNull($margins === null ? null : self::AttributeText($margins, 'header')),
            ComparisonValueHelpers::ParseDoubleValueOrNull($margins === null ? null : self::AttributeText($margins, 'footer')),
            ComparisonValueHelpers::ParseIntValue($pageSetup === null ? null : self::AttributeText($pageSetup, 'paperSize')),
            self::AttributeTextOrEmpty($pageSetup, 'orientation'),
            ComparisonValueHelpers::ParseIntValue($pageSetup === null ? null : self::AttributeText($pageSetup, 'firstPageNumber')),
            ComparisonValueHelpers::ParseIntValue($pageSetup === null ? null : self::AttributeText($pageSetup, 'scale')),
            ComparisonValueHelpers::ParseIntValue($pageSetup === null ? null : self::AttributeText($pageSetup, 'fitToWidth')),
            ComparisonValueHelpers::ParseIntValue($pageSetup === null ? null : self::AttributeText($pageSetup, 'fitToHeight')),
            ComparisonValueHelpers::ParseBoolValue($printOptions === null ? null : self::AttributeText($printOptions, 'gridLines')),
            ComparisonValueHelpers::ParseBoolValue($printOptions === null ? null : self::AttributeText($printOptions, 'headings')),
            ComparisonValueHelpers::ParseBoolValue($printOptions === null ? null : self::AttributeText($printOptions, 'horizontalCentered')),
            ComparisonValueHelpers::ParseBoolValue($printOptions === null ? null : self::AttributeText($printOptions, 'verticalCentered')),
            self::ChildTextOrEmpty($headerFooter, self::MainNs . 'oddHeader'),
            self::ChildTextOrEmpty($headerFooter, self::MainNs . 'oddFooter'),
            $rowBreaks,
            $columnBreaks
        );
    }

    private static function ReadDefinedNames($workbookPath)
    {
        $package = new OpenXmlPackage($workbookPath);
        $workbookDocument = $package->ReadXml('xl/workbook.xml');
        $workbookRoot = $workbookDocument->Root;
        if ($workbookRoot === null) {
            throw new \RuntimeException('Workbook XML is empty.');
        }

        $definedNames = array();
        $definedNamesElement = $workbookRoot->Element(self::MainNs . 'definedNames');
        if ($definedNamesElement !== null) {
            foreach ($definedNamesElement->Elements(self::MainNs . 'definedName') as $definedName) {
                $name = self::AttributeTextOrEmpty($definedName, 'name');
                if ($name === '_xlnm.Print_Area' || $name === '_xlnm.Print_Titles') {
                    continue;
                }

                $definedNames[] = new DefinedNameSnapshot(
                    $name,
                    $definedName->Value,
                    ComparisonValueHelpers::ParseIntValue(self::AttributeText($definedName, 'localSheetId')),
                    ComparisonValueHelpers::ParseBoolValue(self::AttributeText($definedName, 'hidden')),
                    self::AttributeTextOrEmpty($definedName, 'comment')
                );
            }
        }

        return $definedNames;
    }

    private static function AssertDefinedName($actual, $name, $formula, $localSheetIndex, $hidden, $comment)
    {
        AssertEx::Equal($name, $actual->Name);
        AssertEx::Equal($formula, $actual->Formula);
        AssertEx::Equal($localSheetIndex, $actual->LocalSheetIndex);
        AssertEx::Equal($hidden, $actual->Hidden);
        AssertEx::Equal($comment, $actual->Comment);
    }

    private static function AssertSheet($actual, $expectedName, $expectedVisibility)
    {
        AssertEx::Equal($expectedName, $actual->Name);
        AssertEx::Equal($expectedVisibility, $actual->Visibility);
    }

    private static function AssertRow($snapshot, $rowIndex, $height, $hidden)
    {
        AssertEx::True(array_key_exists($rowIndex, $snapshot->Rows), 'Expected row ' . $rowIndex . ' to exist in Open XML snapshot.');
        $row = $snapshot->Rows[$rowIndex];
        AssertEx::Equal($height, $row->Height);
        AssertEx::Equal($hidden, $row->Hidden);
    }

    private static function AssertColumn($snapshot, $minColumnIndex, $maxColumnIndex, $width, $hidden)
    {
        $column = null;
        foreach ($snapshot->Columns as $item) {
            if ($item->MinColumnIndex === $minColumnIndex && $item->MaxColumnIndex === $maxColumnIndex) {
                $column = $item;
                break;
            }
        }

        AssertEx::NotNull($column, 'Expected column range ' . $minColumnIndex . ':' . $maxColumnIndex . ' to exist in Open XML snapshot.');
        AssertEx::Equal($width, $column->Width);
        AssertEx::Equal($hidden, $column->Hidden);
    }

    private static function AssertHyperlink($actual, $area, $address, $location, $screenTip, $display)
    {
        AssertEx::Equal($area, $actual->Area);
        AssertEx::Equal($address, $actual->Address);
        AssertEx::Equal($location, $actual->Location);
        AssertEx::Equal($screenTip, $actual->ScreenTip);
        AssertEx::Equal($display, $actual->Display);
    }

    private static function GetValidationSnapshot($snapshots, $sqref)
    {
        foreach ($snapshots as $snapshot) {
            if ($snapshot->Sqref === $sqref) {
                return $snapshot;
            }
        }

        throw new \RuntimeException('Validation snapshot was not found: ' . $sqref);
    }

    private static function GetConditionalFormattingSnapshot($snapshots, $sqref)
    {
        foreach ($snapshots as $snapshot) {
            if ($snapshot->Sqref === $sqref) {
                return $snapshot;
            }
        }

        throw new \RuntimeException('Conditional formatting snapshot was not found: ' . $sqref);
    }

    private static function AssertValidation($actual, $sqref, $type, $operatorName, $errorStyle, $allowBlank, $showDropDown, $showInputMessage, $showErrorMessage, $promptTitle, $prompt, $errorTitle, $error, $formula1, $formula2)
    {
        AssertEx::Equal($sqref, $actual->Sqref);
        AssertEx::Equal($type, $actual->Type);
        AssertEx::Equal($operatorName, $actual->Operator);
        AssertEx::Equal($errorStyle, $actual->ErrorStyle);
        AssertEx::Equal($allowBlank, $actual->AllowBlank);
        AssertEx::Equal($showDropDown, $actual->ShowDropDown);
        AssertEx::Equal($showInputMessage, $actual->ShowInputMessage);
        AssertEx::Equal($showErrorMessage, $actual->ShowErrorMessage);
        AssertEx::Equal($promptTitle, $actual->PromptTitle);
        AssertEx::Equal($prompt, $actual->Prompt);
        AssertEx::Equal($errorTitle, $actual->ErrorTitle);
        AssertEx::Equal($error, $actual->Error);
        AssertEx::Equal($formula1, $actual->Formula1);
        AssertEx::Equal($formula2, $actual->Formula2);
    }

    private static function AssertConditionalFormatting($actual, $sqref, $ruleCount)
    {
        AssertEx::Equal($sqref, $actual->Sqref);
        AssertEx::Equal($ruleCount, count($actual->Rules));
    }

    private static function ResolveVisibility($state)
    {
        switch ($state) {
            case 'hidden':
                return 'Hidden';
            case 'veryHidden':
                return 'VeryHidden';
            default:
                return 'Visible';
        }
    }

    private static function UnqualifyDefinedNameValue($value)
    {
        $trimmed = trim($value);
        $exclamationIndex = strrpos($trimmed, '!');
        return $exclamationIndex === false ? $trimmed : substr($trimmed, $exclamationIndex + 1);
    }

    private static function AttributeText($element, $name)
    {
        return ComparisonValueHelpers::AttributeText($element, $name);
    }

    private static function AttributeTextOrEmpty($element, $name)
    {
        $value = ComparisonValueHelpers::AttributeText($element, $name);
        return $value === null ? '' : $value;
    }

    private static function ChildTextOrEmpty($element, $name)
    {
        if ($element === null) {
            return '';
        }

        $child = $element->Element($name);
        return $child === null ? '' : $child->Value;
    }

    private static function ChildAttributeText($element, $childName, $attributeName)
    {
        if ($element === null) {
            return '';
        }

        $child = $element->Element($childName);
        return $child === null ? '' : self::AttributeTextOrEmpty($child, $attributeName);
    }

    private static function ParseUInt($value)
    {
        if ($value === null || preg_match('/^\s*\d+\s*$/', $value) !== 1) {
            return 0;
        }

        return (int) $value;
    }

    private static function ResolvePart($relationships, $typeFragment)
    {
        foreach ($relationships as $relationship) {
            if (strpos($relationship['Type'], $typeFragment) === false) {
                continue;
            }

            $target = $relationship['ResolvedTarget'];
            if ($target !== null && $target !== '') {
                return ltrim($target, '/');
            }
        }

        return null;
    }
}

final class TableTestMethods
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';

    public static function TableWorkbookMatchesOpenXmlSnapshot()
    {
        $temp = new TemporaryDirectory('table-openxml');
        $path = $temp->GetPath('table-test.xlsx');

        $workbook = ListObjectScenarioFactory::CreateTableWorkbook();
        $workbook->Save($path);

        $package = new OpenXmlPackage($path);
        $worksheetPart = ProgramTestSupport::FirstWorksheetPart($package);
        $tableParts = ProgramTestSupport::RelationshipsByType($package, $worksheetPart, '/table');
        AssertEx::Equal(1, count($tableParts));

        $tableDocument = $package->ReadXml($tableParts[0]['ResolvedTarget']);
        $root = $tableDocument->Root;
        AssertEx::NotNull($root, 'Table root must not be null.');

        AssertEx::Equal('Table1', ProgramTestSupport::AttributeOrNull($root, 'displayName'));
        AssertEx::Equal('A1:C5', ProgramTestSupport::AttributeOrNull($root, 'ref'));
        AssertEx::Equal('1', ProgramTestSupport::AttributeOrDefault($root, 'totalsRowCount', '0'));

        $tableColumnsElement = $root->Element(self::MainNs . 'tableColumns');
        AssertEx::NotNull($tableColumnsElement, 'tableColumns element must be present.');
        AssertEx::Equal('3', ProgramTestSupport::AttributeOrNull($tableColumnsElement, 'count'));

        $columns = $tableColumnsElement->Elements(self::MainNs . 'tableColumn');
        AssertEx::Equal(3, count($columns));
        AssertEx::Equal('Product', ProgramTestSupport::AttributeOrNull($columns[0], 'name'));
        AssertEx::Equal('Region', ProgramTestSupport::AttributeOrNull($columns[1], 'name'));
        AssertEx::Equal('Revenue', ProgramTestSupport::AttributeOrNull($columns[2], 'name'));
        AssertEx::Equal('sum', ProgramTestSupport::AttributeOrNull($columns[2], 'totalsRowFunction'));
        AssertEx::Equal('Total', ProgramTestSupport::AttributeOrNull($columns[0], 'totalsRowLabel'));

        $styleInfo = $root->Element(self::MainNs . 'tableStyleInfo');
        AssertEx::NotNull($styleInfo, 'tableStyleInfo element must be present.');
        AssertEx::Equal('TableStyleMedium2', ProgramTestSupport::AttributeOrNull($styleInfo, 'name'));
        AssertEx::Equal('1', ProgramTestSupport::AttributeOrNull($styleInfo, 'showRowStripes'));
    }

    public static function TableWorkbookRoundtrip()
    {
        $temp = new TemporaryDirectory('table-roundtrip');
        $path = $temp->GetPath('table-roundtrip.xlsx');

        $workbook = ListObjectScenarioFactory::CreateTableWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        ListObjectScenarioFactory::AssertTableWorkbook($loaded);
    }
}

final class PictureTestMethods
{
    const XdrNs = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';
    const ANs = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    public static function PictureWorkbookMatchesOpenXmlSnapshot()
    {
        $temp = new TemporaryDirectory('picture-openxml');
        $path = $temp->GetPath('picture-test.xlsx');

        $workbook = PictureScenarioFactory::CreatePictureWorkbook();
        $workbook->Save($path);

        $package = new OpenXmlPackage($path);
        $worksheetPart = ProgramTestSupport::FirstWorksheetPart($package);
        $drawingParts = ProgramTestSupport::RelationshipsByType($package, $worksheetPart, '/drawing');
        AssertEx::True(count($drawingParts) > 0, 'DrawingsPart must be present on the worksheet.');

        $drawingDocument = $package->ReadXml($drawingParts[0]['ResolvedTarget']);
        $root = $drawingDocument->Root;
        AssertEx::NotNull($root, 'Drawing root must not be null.');

        $anchors = $root->Elements(self::XdrNs . 'twoCellAnchor');
        AssertEx::Equal(2, count($anchors));

        $anchor0 = $anchors[0];
        $from0 = $anchor0->Element(self::XdrNs . 'from');
        AssertEx::NotNull($from0, 'from element must be present on first anchor.');
        AssertEx::Equal('0', ProgramTestSupport::ChildTextOrEmpty($from0, self::XdrNs . 'col'));
        AssertEx::Equal('0', ProgramTestSupport::ChildTextOrEmpty($from0, self::XdrNs . 'row'));
        $to0 = $anchor0->Element(self::XdrNs . 'to');
        AssertEx::NotNull($to0, 'to element must be present on first anchor.');
        AssertEx::Equal('5', ProgramTestSupport::ChildTextOrEmpty($to0, self::XdrNs . 'col'));
        AssertEx::Equal('8', ProgramTestSupport::ChildTextOrEmpty($to0, self::XdrNs . 'row'));
        AssertEx::NotNull($anchor0->Element(self::XdrNs . 'pic'), 'pic element must be present on first anchor.');

        $anchor1 = $anchors[1];
        $from1 = $anchor1->Element(self::XdrNs . 'from');
        AssertEx::NotNull($from1, 'from element must be present on second anchor.');
        AssertEx::Equal('0', ProgramTestSupport::ChildTextOrEmpty($from1, self::XdrNs . 'col'));
        AssertEx::Equal('10', ProgramTestSupport::ChildTextOrEmpty($from1, self::XdrNs . 'row'));
        AssertEx::NotNull($anchor1->Element(self::XdrNs . 'pic'), 'pic element must be present on second anchor.');

        $imageParts = ProgramTestSupport::RelationshipsByType($package, $drawingParts[0]['ResolvedTarget'], '/image');
        AssertEx::Equal(2, count($imageParts));
    }

    public static function PictureWorkbookRoundtrip()
    {
        $temp = new TemporaryDirectory('picture-roundtrip');
        $path = $temp->GetPath('picture-roundtrip.xlsx');

        $workbook = PictureScenarioFactory::CreatePictureWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        PictureScenarioFactory::AssertPictureWorkbook($loaded);
    }
}

final class ShapeTestMethods
{
    const XdrNs = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';
    const ANs = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    public static function ShapeWorkbookMatchesOpenXmlSnapshot()
    {
        $temp = new TemporaryDirectory('shape-openxml');
        $path = $temp->GetPath('shape-test.xlsx');

        $workbook = ShapeScenarioFactory::CreateShapeWorkbook();
        $workbook->Save($path);

        $package = new OpenXmlPackage($path);
        $worksheetPart = ProgramTestSupport::FirstWorksheetPart($package);
        $drawingParts = ProgramTestSupport::RelationshipsByType($package, $worksheetPart, '/drawing');
        AssertEx::True(count($drawingParts) > 0, 'DrawingsPart must be present on the worksheet.');

        $drawingDocument = $package->ReadXml($drawingParts[0]['ResolvedTarget']);
        $root = $drawingDocument->Root;
        AssertEx::NotNull($root, 'Drawing root must not be null.');

        $anchors = $root->Elements(self::XdrNs . 'twoCellAnchor');
        AssertEx::Equal(6, count($anchors));

        $sp0 = $anchors[0]->Element(self::XdrNs . 'sp');
        AssertEx::NotNull($sp0, 'First anchor must contain an xdr:sp element.');
        $nvSpPr0 = $sp0->Element(self::XdrNs . 'nvSpPr');
        $cNvPr0 = $nvSpPr0 === null ? null : $nvSpPr0->Element(self::XdrNs . 'cNvPr');
        AssertEx::Equal('MyRectangle', $cNvPr0 === null ? '' : ProgramTestSupport::AttributeOrEmpty($cNvPr0, 'name'));
        $spPr0 = $sp0->Element(self::XdrNs . 'spPr');
        $prstGeom0 = $spPr0 === null ? null : $spPr0->Element(self::ANs . 'prstGeom');
        AssertEx::Equal('rect', $prstGeom0 === null ? '' : ProgramTestSupport::AttributeOrEmpty($prstGeom0, 'prst'));

        $sp1 = $anchors[1]->Element(self::XdrNs . 'sp');
        AssertEx::NotNull($sp1, 'Second anchor must contain an xdr:sp element.');
        $spPr1 = $sp1->Element(self::XdrNs . 'spPr');
        $prstGeom1 = $spPr1 === null ? null : $spPr1->Element(self::ANs . 'prstGeom');
        AssertEx::Equal('rightArrow', $prstGeom1 === null ? '' : ProgramTestSupport::AttributeOrEmpty($prstGeom1, 'prst'));

        $sp2 = $anchors[2]->Element(self::XdrNs . 'sp');
        AssertEx::NotNull($sp2, 'Third anchor must contain an xdr:sp element.');
        $spPr2 = $sp2->Element(self::XdrNs . 'spPr');
        $prstGeom2 = $spPr2 === null ? null : $spPr2->Element(self::ANs . 'prstGeom');
        AssertEx::Equal('ellipse', $prstGeom2 === null ? '' : ProgramTestSupport::AttributeOrEmpty($prstGeom2, 'prst'));

        for ($index = 0; $index < 6; $index++) {
            $sp = $anchors[$index]->Element(self::XdrNs . 'sp');
            AssertEx::NotNull($sp, 'Anchor ' . $index . ' must have sp element.');
            AssertEx::NotNull($sp->Element(self::XdrNs . 'style'), 'Shape ' . $index . ' must have xdr:style element for visibility.');
        }

        $imageParts = ProgramTestSupport::RelationshipsByType($package, $drawingParts[0]['ResolvedTarget'], '/image');
        AssertEx::Equal(0, count($imageParts));
    }

    public static function ShapeWorkbookRoundtrip()
    {
        $temp = new TemporaryDirectory('shape-roundtrip');
        $path = $temp->GetPath('shape-roundtrip.xlsx');

        $workbook = ShapeScenarioFactory::CreateShapeWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        ShapeScenarioFactory::AssertShapeWorkbook($loaded);
    }
}

final class ProgramTestSupport
{
    public static function FirstWorksheetPart($package)
    {
        $relationships = $package->ReadRelationships('xl/workbook.xml');
        foreach ($relationships as $relationship) {
            if (strpos($relationship['Type'], '/worksheet') === false) {
                continue;
            }

            $target = $relationship['ResolvedTarget'];
            if ($target !== null && $target !== '') {
                return ltrim($target, '/');
            }
        }

        throw new \RuntimeException('Worksheet part was not found.');
    }

    public static function RelationshipsByType($package, $sourcePart, $typeFragment)
    {
        $result = array();
        foreach ($package->ReadRelationships($sourcePart) as $relationship) {
            if (strpos($relationship['Type'], $typeFragment) === false) {
                continue;
            }

            $result[] = $relationship;
        }

        return $result;
    }

    public static function AttributeOrNull($element, $name)
    {
        return ComparisonValueHelpers::AttributeText($element, $name);
    }

    public static function AttributeOrEmpty($element, $name)
    {
        $value = ComparisonValueHelpers::AttributeText($element, $name);
        return $value === null ? '' : $value;
    }

    public static function AttributeOrDefault($element, $name, $default)
    {
        $value = ComparisonValueHelpers::AttributeText($element, $name);
        return $value === null ? $default : $value;
    }

    public static function ChildTextOrEmpty($element, $name)
    {
        $child = $element->Element($name);
        return $child === null ? '' : $child->Value;
    }
}

Program::Main();
