<?php

namespace Aspose\Cells_FOSS\Testing;

use Aspose\Cells_FOSS\BorderStyleType;
use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\HorizontalAlignmentType;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\PageOrientationType;
use Aspose\Cells_FOSS\PaperSizeType;
use Aspose\Cells_FOSS\ValidationAlertType;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\VerticalAlignmentType;
use Aspose\Cells_FOSS\VisibilityType;
use Aspose\Cells_FOSS\Workbook;

final class TestCase
{
    public $Name;
    public $Body;

    public function __construct($name, $body)
    {
        $this->Name = $name;
        $this->Body = $body;
    }
}

final class TestRunner
{
    public static function Run($suiteName, array $tests)
    {
        echo 'Running ' . $suiteName . ' (' . count($tests) . ' tests)' . PHP_EOL;
        $failed = 0;

        foreach ($tests as $test) {
            try {
                call_user_func($test->Body);
                echo 'PASS ' . $test->Name . PHP_EOL;
            } catch (\Throwable $exception) {
                $failed++;
                echo 'FAIL ' . $test->Name . PHP_EOL;
                echo $exception . PHP_EOL;
            }
        }

        echo $suiteName . ': ' . (count($tests) - $failed) . ' passed, ' . $failed . ' failed' . PHP_EOL;
        return $failed === 0 ? 0 : 1;
    }
}

final class AssertEx
{
    public static function True($condition, $message = null)
    {
        if (!$condition) {
            throw new \RuntimeException($message === null ? 'Expected condition to be true.' : $message);
        }
    }

    public static function False($condition, $message = null)
    {
        if ($condition) {
            throw new \RuntimeException($message === null ? 'Expected condition to be false.' : $message);
        }
    }

    public static function Equal($expected, $actual, $message = null)
    {
        if (!($expected == $actual)) {
            throw new \RuntimeException($message === null
                ? "Expected '" . self::describe($expected) . "', got '" . self::describe($actual) . "'."
                : $message);
        }
    }

    public static function NotNull($value, $message = null)
    {
        if ($value === null) {
            throw new \RuntimeException($message === null ? 'Expected value to be non-null.' : $message);
        }
    }

    public static function Null($value, $message = null)
    {
        if ($value !== null) {
            throw new \RuntimeException($message === null
                ? "Expected value to be null, got '" . self::describe($value) . "'."
                : $message);
        }
    }

    public static function Contains($expectedSubstring, $actual, $message = null)
    {
        if (strpos($actual, $expectedSubstring) === false) {
            throw new \RuntimeException($message === null
                ? "Expected '" . $actual . "' to contain '" . $expectedSubstring . "'."
                : $message);
        }
    }

    public static function Throws($exceptionClass, callable $action, $message = null)
    {
        try {
            call_user_func($action);
        } catch (\Throwable $exception) {
            if ($exception instanceof $exceptionClass) {
                return $exception;
            }

            throw new \RuntimeException($message === null
                ? 'Expected ' . $exceptionClass . ', got ' . get_class($exception) . '.'
                : $message);
        }

        throw new \RuntimeException($message === null
            ? 'Expected ' . $exceptionClass . ' to be thrown.'
            : $message);
    }

    private static function describe($value)
    {
        if ($value === null) {
            return 'null';
        }

        if (is_bool($value)) {
            return $value ? 'true' : 'false';
        }

        if (is_array($value)) {
            return 'array';
        }

        if (is_object($value)) {
            if (method_exists($value, '__toString')) {
                return (string) $value;
            }

            return get_class($value);
        }

        return (string) $value;
    }
}

final class TemporaryDirectory
{
    public $RootPath;

    public function __construct($suiteName)
    {
        $repositoryRoot = self::ResolveRepositoryRoot();
        $this->RootPath = $repositoryRoot . DIRECTORY_SEPARATOR . 'output' . DIRECTORY_SEPARATOR . $suiteName . DIRECTORY_SEPARATOR . self::newGuid();
        if (!is_dir($this->RootPath)) {
            mkdir($this->RootPath, 0777, true);
        }
    }

    public function GetPath($fileName)
    {
        return $this->RootPath . DIRECTORY_SEPARATOR . $fileName;
    }

    public function Dispose()
    {
        // Preserve generated artifacts under the repo-local output folder for inspection.
    }

    private static function ResolveRepositoryRoot()
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== '') {
            if (is_file($directory . DIRECTORY_SEPARATOR . 'composer.json')) {
                return $directory;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }

            $directory = $parent;
        }

        return getcwd();
    }

    private static function newGuid()
    {
        return bin2hex(random_bytes(16));
    }
}

final class WorkbookScenarioFactory
{
    public static function CreateMixedCellWorkbook($useDate1904 = false)
    {
        $workbook = new Workbook();
        $workbook->Settings->Date1904 = $useDate1904;

        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Data';
        $sheet->Cells['A1']->PutValue('Hello');
        $sheet->Cells['B1']->PutValue(123);
        $sheet->Cells['C1']->PutValue(true);
        $sheet->Cells['D1']->PutValue(12.5);
        $sheet->Cells['E1']->PutValue(6.02214076E+23);
        $sheet->Cells['F1']->PutValue(new \DateTime('2024-05-06 07:08:09'));
        $sheet->Cells['G1']->PutValue(20);
        $sheet->Cells['G1']->Formula = '=B1*2';
        return $workbook;
    }
}

final class WorksheetScenarioFactory
{
    public static function CreateWorksheetSettingsWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Layout';
        $sheet->VisibilityType = VisibilityType::Hidden;
        $sheet->TabColor = Color::FromArgb(255, 34, 68, 102);
        $sheet->ShowGridlines = false;
        $sheet->ShowRowColumnHeaders = false;
        $sheet->ShowZeros = false;
        $sheet->RightToLeft = true;
        $sheet->Zoom = 85;
        $sheet->Protect();
        $sheet->Protection->Objects = true;
        $sheet->Protection->Scenarios = true;
        $sheet->Protection->FormatCells = true;
        $sheet->Protection->InsertRows = true;
        $sheet->Protection->AutoFilter = true;
        $sheet->Protection->SelectLockedCells = true;
        $sheet->Protection->SelectUnlockedCells = true;

        $sheet->Cells['A1']->PutValue('Merged');
        $sheet->Cells['C4']->PutValue(99);
        $sheet->Cells->Rows[1]->Height = 22.5;
        $sheet->Cells->Rows[3]->IsHidden = true;
        $sheet->Cells->Columns[0]->Width = 18.25;
        $sheet->Cells->Columns[2]->IsHidden = true;
        $sheet->Cells->Merge(0, 0, 2, 2);

        $visibleSheetIndex = $workbook->Worksheets->Add();
        $visibleSheet = $workbook->Worksheets[$visibleSheetIndex];
        $visibleSheet->Name = 'Visible';
        $visibleSheet->Cells['A1']->PutValue('Visible');
        $workbook->Worksheets->ActiveSheetName = 'Visible';

        return $workbook;
    }

    public static function AssertWorksheetSettings(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets['Layout'];
        AssertEx::NotNull($sheet, "Expected the worksheet settings scenario to contain the 'Layout' sheet.");
        AssertEx::Equal(VisibilityType::Hidden, $sheet->VisibilityType);
        AssertEx::Equal(Color::FromArgb(255, 34, 68, 102), $sheet->TabColor);
        AssertEx::False($sheet->ShowGridlines);
        AssertEx::False($sheet->ShowRowColumnHeaders);
        AssertEx::False($sheet->ShowZeros);
        AssertEx::True($sheet->RightToLeft);
        AssertEx::Equal(85, $sheet->Zoom);
        AssertEx::True($sheet->Protection->IsProtected);
        AssertEx::True($sheet->Protection->Objects);
        AssertEx::True($sheet->Protection->Scenarios);
        AssertEx::True($sheet->Protection->FormatCells);
        AssertEx::True($sheet->Protection->InsertRows);
        AssertEx::True($sheet->Protection->AutoFilter);
        AssertEx::True($sheet->Protection->SelectLockedCells);
        AssertEx::True($sheet->Protection->SelectUnlockedCells);
        AssertEx::Equal('Merged', $sheet->Cells['A1']->StringValue);
        AssertEx::Equal(99, (int) $sheet->Cells['C4']->Value);
        AssertEx::Equal(22.5, $sheet->Cells->Rows[1]->Height ?? 0.0);
        AssertEx::True($sheet->Cells->Rows[3]->IsHidden);
        AssertEx::Equal(18.25, $sheet->Cells->Columns[0]->Width ?? 0.0);
        AssertEx::True($sheet->Cells->Columns[2]->IsHidden);
        AssertEx::Equal(1, count($sheet->Cells->MergedCells));
        AssertEx::Equal(0, $sheet->Cells->MergedCells[0]->StartRow);
        AssertEx::Equal(0, $sheet->Cells->MergedCells[0]->StartColumn);
        AssertEx::Equal(2, $sheet->Cells->MergedCells[0]->EndRow - $sheet->Cells->MergedCells[0]->StartRow + 1);
        AssertEx::Equal(2, $sheet->Cells->MergedCells[0]->EndColumn - $sheet->Cells->MergedCells[0]->StartColumn + 1);
    }

    public static function AssertWorksheetSettingsScenarioHasVisibleSheet(Workbook $workbook)
    {
        AssertEx::True($workbook->Worksheets->Count >= 2);
        $visibleSheet = $workbook->Worksheets['Visible'];
        AssertEx::NotNull($visibleSheet, 'Expected the worksheet settings scenario to contain a visible sheet.');
        AssertEx::Equal(VisibilityType::Visible, $visibleSheet->VisibilityType);
        AssertEx::Equal('Visible', $visibleSheet->Cells['A1']->StringValue);
        AssertEx::Equal('Visible', $workbook->Worksheets->ActiveSheetName);
    }
}

final class PageSetupScenarioFactory
{
    public static function CreatePageSetupWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Print Sheet';
        $sheet->Cells['A1']->PutValue('Title');
        $sheet->Cells['C10']->PutValue(42);

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
        $pageSetup->PrintArea = '$A$1:$C$10';
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

        return $workbook;
    }

    public static function AssertPageSetup(Workbook $workbook)
    {
        $pageSetup = $workbook->Worksheets[0]->PageSetup;
        AssertEx::Equal(0.25, $pageSetup->LeftMarginInch);
        AssertEx::Equal(0.4, $pageSetup->RightMarginInch);
        AssertEx::Equal(0.5, $pageSetup->TopMarginInch);
        AssertEx::Equal(0.6, $pageSetup->BottomMarginInch);
        AssertEx::Equal(0.2, $pageSetup->HeaderMarginInch);
        AssertEx::Equal(0.22, $pageSetup->FooterMarginInch);
        AssertEx::Equal(PageOrientationType::Landscape, $pageSetup->Orientation);
        AssertEx::Equal(PaperSizeType::PaperA4, $pageSetup->PaperSize);
        AssertEx::Equal(3, $pageSetup->FirstPageNumber ?? 0);
        AssertEx::Equal(95, $pageSetup->Scale ?? 0);
        AssertEx::Equal(1, $pageSetup->FitToPagesWide ?? 0);
        AssertEx::Equal(2, $pageSetup->FitToPagesTall ?? 0);
        AssertEx::Equal('$A$1:$C$10', $pageSetup->PrintArea);
        AssertEx::Equal('$1:$2', $pageSetup->PrintTitleRows);
        AssertEx::Equal('$A:$B', $pageSetup->PrintTitleColumns);
        AssertEx::Equal('Left Header', $pageSetup->LeftHeader);
        AssertEx::Equal('Center Header', $pageSetup->CenterHeader);
        AssertEx::Equal('Right Header', $pageSetup->RightHeader);
        AssertEx::Equal('Left Footer', $pageSetup->LeftFooter);
        AssertEx::Equal('Center Footer', $pageSetup->CenterFooter);
        AssertEx::Equal('Right Footer', $pageSetup->RightFooter);
        AssertEx::True($pageSetup->PrintGridlines);
        AssertEx::True($pageSetup->PrintHeadings);
        AssertEx::True($pageSetup->CenterHorizontally);
        AssertEx::True($pageSetup->CenterVertically);
        AssertEx::Equal(2, count($pageSetup->HorizontalPageBreaks));
        AssertEx::Equal(4, $pageSetup->HorizontalPageBreaks[0]);
        AssertEx::Equal(7, $pageSetup->HorizontalPageBreaks[1]);
        AssertEx::Equal(1, count($pageSetup->VerticalPageBreaks));
        AssertEx::Equal(2, $pageSetup->VerticalPageBreaks[0]);
        AssertEx::Equal('Title', $workbook->Worksheets[0]->Cells['A1']->StringValue);
        AssertEx::Equal('42', $workbook->Worksheets[0]->Cells['C10']->StringValue);
    }
}

final class HyperlinkScenarioFactory
{
    public static function CreateHyperlinkWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Links';
        $workbook->Worksheets->Add('Target Sheet');
        $workbook->Worksheets[1]->Cells[[2, 2]]->PutValue('Target');

        $sheet->Cells['A1']->PutValue('Docs');
        $external = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('A1', 1, 1, 'https://example.com/docs?q=1')];
        $external->TextToDisplay = 'Docs';
        $external->ScreenTip = 'External docs';

        $sheet->Cells['B2']->PutValue('Jump');
        $internalLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('B2', 1, 1, "'Target Sheet'!C3")];
        $internalLink->TextToDisplay = 'Jump';
        $internalLink->ScreenTip = 'Jump to target';

        $rangeLink = $sheet->Hyperlinks[$sheet->Hyperlinks->Add('C4', 2, 2, 'mailto:test@example.com')];
        $rangeLink->TextToDisplay = 'Mail';
        $rangeLink->ScreenTip = 'Send mail';

        return $workbook;
    }

    public static function AssertHyperlinks(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets[0];
        AssertEx::Equal(3, $sheet->Hyperlinks->Count);

        $external = $sheet->Hyperlinks[0];
        AssertEx::Equal('A1', $external->Area);
        AssertEx::Equal('https://example.com/docs?q=1', $external->Address);
        AssertEx::Equal('External docs', $external->ScreenTip);
        AssertEx::Equal('Docs', $external->TextToDisplay);

        $internalLink = $sheet->Hyperlinks[1];
        AssertEx::Equal('B2', $internalLink->Area);
        AssertEx::Equal("'Target Sheet'!C3", $internalLink->Address);
        AssertEx::Equal('Jump to target', $internalLink->ScreenTip);
        AssertEx::Equal('Jump', $internalLink->TextToDisplay);

        $rangeLink = $sheet->Hyperlinks[2];
        AssertEx::Equal('C4:D5', $rangeLink->Area);
        AssertEx::Equal('mailto:test@example.com', $rangeLink->Address);
        AssertEx::Equal('Send mail', $rangeLink->ScreenTip);
        AssertEx::Equal('Mail', $rangeLink->TextToDisplay);
    }
}

final class ZipPackageHelper
{
    public static function ReadEntryText($packagePath, $entryPath)
    {
        $archive = new \ZipArchive();
        if ($archive->open($packagePath) !== true) {
            throw new \RuntimeException("The package '" . $packagePath . "' could not be opened.");
        }

        $content = $archive->getFromName(self::NormalizeEntryPath($entryPath));
        $archive->close();
        if ($content === false) {
            throw new \RuntimeException("Missing zip entry '" . $entryPath . "'.");
        }

        return $content;
    }

    public static function EntryExists($packagePath, $entryPath)
    {
        $archive = new \ZipArchive();
        if ($archive->open($packagePath) !== true) {
            throw new \RuntimeException("The package '" . $packagePath . "' could not be opened.");
        }

        $exists = $archive->locateName(self::NormalizeEntryPath($entryPath)) !== false;
        $archive->close();
        return $exists;
    }

    public static function RewriteEntryText($packagePath, $entryPath, callable $rewrite)
    {
        $archive = new \ZipArchive();
        if ($archive->open($packagePath) !== true) {
            throw new \RuntimeException("The package '" . $packagePath . "' could not be opened.");
        }

        $normalizedPath = self::NormalizeEntryPath($entryPath);
        $content = $archive->getFromName($normalizedPath);
        if ($content === false) {
            $archive->close();
            throw new \RuntimeException("Missing zip entry '" . $entryPath . "'.");
        }

        $archive->deleteName($normalizedPath);
        $archive->addFromString($normalizedPath, call_user_func($rewrite, $content));
        $archive->close();
    }

    public static function RewriteXmlEntry($packagePath, $entryPath, callable $mutate)
    {
        self::RewriteEntryText($packagePath, $entryPath, function ($content) use ($mutate) {
            $document = XDocument::Load($content);
            call_user_func($mutate, $document);
            return $document->save();
        });
    }

    public static function MoveEntry($packagePath, $entryPath, $destinationPath)
    {
        $archive = new \ZipArchive();
        if ($archive->open($packagePath) !== true) {
            throw new \RuntimeException("The package '" . $packagePath . "' could not be opened.");
        }

        $normalizedPath = self::NormalizeEntryPath($entryPath);
        $content = $archive->getFromName($normalizedPath);
        if ($content === false) {
            $archive->close();
            throw new \RuntimeException("Missing zip entry '" . $entryPath . "'.");
        }

        $normalizedDestination = self::NormalizeEntryPath($destinationPath);
        if ($archive->locateName($normalizedDestination) !== false) {
            $archive->deleteName($normalizedDestination);
        }

        $archive->deleteName($normalizedPath);
        $archive->addFromString($normalizedDestination, $content);
        $archive->close();
    }

    public static function DeleteEntry($packagePath, $entryPath)
    {
        $archive = new \ZipArchive();
        if ($archive->open($packagePath) !== true) {
            throw new \RuntimeException("The package '" . $packagePath . "' could not be opened.");
        }

        $normalizedPath = self::NormalizeEntryPath($entryPath);
        if ($archive->locateName($normalizedPath) !== false) {
            $archive->deleteName($normalizedPath);
        }

        $archive->close();
    }

    public static function CreatePackage($packagePath, array $entries)
    {
        $archive = new \ZipArchive();
        if ($archive->open($packagePath, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) !== true) {
            throw new \RuntimeException("The package '" . $packagePath . "' could not be created.");
        }

        foreach ($entries as $key => $value) {
            $archive->addFromString(self::NormalizeEntryPath($key), $value);
        }

        $archive->close();
    }

    private static function NormalizeEntryPath($entryPath)
    {
        return ltrim(str_replace('\\', '/', $entryPath), '/');
    }
}

final class StyleScenarioFactory
{
    private static $FillPatternShowcase = array(
        FillPattern::Solid,
        FillPattern::MediumGray,
        FillPattern::DarkGray,
        FillPattern::Gray125,
        FillPattern::Gray0625,
        FillPattern::DarkHorizontal,
        FillPattern::DarkVertical,
        FillPattern::DarkDown,
        FillPattern::DarkUp,
        FillPattern::DarkGrid,
        FillPattern::DarkTrellis,
        FillPattern::LightHorizontal,
        FillPattern::LightVertical,
        FillPattern::LightDown,
        FillPattern::LightUp,
        FillPattern::LightGrid,
        FillPattern::LightTrellis,
    );

    private static $BorderStyleShowcase = array(
        BorderStyleType::Thin,
        BorderStyleType::Medium,
        BorderStyleType::Thick,
        BorderStyleType::Dotted,
        BorderStyleType::Dashed,
        BorderStyleType::Double,
        BorderStyleType::Hair,
        BorderStyleType::MediumDashed,
        BorderStyleType::DashDot,
        BorderStyleType::MediumDashDot,
        BorderStyleType::DashDotDot,
        BorderStyleType::MediumDashDotDot,
        BorderStyleType::SlantedDashDot,
    );

    private static $HorizontalAlignmentShowcase = array(
        HorizontalAlignmentType::General,
        HorizontalAlignmentType::Left,
        HorizontalAlignmentType::Center,
        HorizontalAlignmentType::Right,
        HorizontalAlignmentType::Fill,
        HorizontalAlignmentType::Justify,
        HorizontalAlignmentType::CenterContinuous,
        HorizontalAlignmentType::Distributed,
    );

    private static $VerticalAlignmentShowcase = array(
        VerticalAlignmentType::Bottom,
        VerticalAlignmentType::Center,
        VerticalAlignmentType::Top,
        VerticalAlignmentType::Justify,
        VerticalAlignmentType::Distributed,
    );

    private static $NumberFormatShowcase = array(
        '0',
        '0.00',
        '#,##0.00',
        '0%',
        '0.00E+00',
        '# ?/?',
        '[$-409]#,##0.00',
    );

    private static $RotationShowcase = array(
        0,
        45,
        90,
        135,
        180,
        255,
    );

    private static $ReadingOrderShowcase = array(
        0,
        1,
        2,
    );

    private static $FontSizeShowcase = array(
        8.0,
        10.0,
        12.0,
        14.0,
        18.0,
        24.0,
    );

    private static $IndentLevelShowcase = array(
        0,
        1,
        2,
        3,
        4,
    );

    const ShowcaseStartColumn = 5;
    const FillPatternRow = 0;
    const BorderStyleRow = 1;
    const HorizontalAlignmentRow = 2;
    const VerticalAlignmentRow = 3;
    const FontRow = 4;
    const NumberFormatRow = 5;
    const BorderSidesRow = 6;
    const RotationRow = 7;
    const ReadingOrderRow = 8;
    const ProtectionRow = 9;
    const FontSizeRow = 10;
    const IndentRow = 11;
    const WrapShrinkRow = 12;

    public static function CreateStyledWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Styled';

        $primaryCell = $sheet->Cells['A1'];
        $primaryCell->PutValue(1234.567);
        $primaryStyle = $primaryCell->GetStyle();
        self::ApplyPrimaryStyle($primaryStyle);
        $primaryCell->SetStyle($primaryStyle);

        $customCell = $sheet->Cells['B2'];
        $customStyle = $customCell->GetStyle();
        self::ApplyCustomNumberStyle($customStyle);
        $customCell->SetStyle($customStyle);

        $dateCell = $sheet->Cells['C3'];
        $dateCell->PutValue(new \DateTime('2024-05-06 07:08:09'));
        $dateStyle = $dateCell->GetStyle();
        self::ApplyDateStyle($dateStyle);
        $dateCell->SetStyle($dateStyle);

        $verticalTextCell = $sheet->Cells['D4'];
        $verticalTextCell->PutValue('Vertical');
        $verticalTextStyle = $verticalTextCell->GetStyle();
        self::ApplyVerticalTextStyle($verticalTextStyle);
        $verticalTextCell->SetStyle($verticalTextStyle);

        self::AddShowcaseLabel($sheet, self::FillPatternRow, 'Fill Patterns');
        self::AddShowcaseLabel($sheet, self::BorderStyleRow, 'Border Styles');
        self::AddShowcaseLabel($sheet, self::HorizontalAlignmentRow, 'Horizontal Align');
        self::AddShowcaseLabel($sheet, self::VerticalAlignmentRow, 'Vertical Align');
        self::AddShowcaseLabel($sheet, self::FontRow, 'Font Settings');
        self::AddShowcaseLabel($sheet, self::NumberFormatRow, 'Number Formats');
        self::AddShowcaseLabel($sheet, self::BorderSidesRow, 'Border Sides');
        self::AddShowcaseLabel($sheet, self::RotationRow, 'Text Rotation');
        self::AddShowcaseLabel($sheet, self::ReadingOrderRow, 'Reading Order');
        self::AddShowcaseLabel($sheet, self::ProtectionRow, 'Protection');
        self::AddShowcaseLabel($sheet, self::FontSizeRow, 'Font Sizes');
        self::AddShowcaseLabel($sheet, self::IndentRow, 'Indent Levels');
        self::AddShowcaseLabel($sheet, self::WrapShrinkRow, 'Wrap And Shrink');

        self::AddFillPatternShowcase($sheet);
        self::AddBorderStyleShowcase($sheet);
        self::AddHorizontalAlignmentShowcase($sheet);
        self::AddVerticalAlignmentShowcase($sheet);
        self::AddFontShowcase($sheet);
        self::AddNumberFormatShowcase($sheet);
        self::AddBorderSideShowcase($sheet);
        self::AddRotationShowcase($sheet);
        self::AddReadingOrderShowcase($sheet);
        self::AddProtectionShowcase($sheet);
        self::AddFontSizeShowcase($sheet);
        self::AddIndentShowcase($sheet);
        self::AddWrapShrinkShowcase($sheet);

        return $workbook;
    }

    public static function ApplyPrimaryStyle($style)
    {
        $style->Font->Name = 'Arial';
        $style->Font->Size = 14.5;
        $style->Font->IsBold = true;
        $style->Font->IsItalic = true;
        $style->Font->Underline = FontUnderlineType::Single;
        $style->Font->IsStrikeout = true;
        $style->Font->Color = Color::FromArgb(255, 17, 34, 51);
        $style->Pattern = FillPattern::LightGrid;
        $style->ForegroundColor = Color::FromArgb(255, 210, 220, 30);
        $style->BackgroundColor = Color::FromArgb(255, 12, 45, 78);
        $style->Borders->Left->LineStyle = BorderStyleType::Dotted;
        $style->Borders->Left->Color = Color::FromArgb(255, 200, 0, 0);
        $style->Borders->Right->LineStyle = BorderStyleType::MediumDashDot;
        $style->Borders->Right->Color = Color::FromArgb(255, 240, 120, 0);
        $style->Borders->Top->LineStyle = BorderStyleType::Double;
        $style->Borders->Top->Color = Color::FromArgb(255, 0, 0, 255);
        $style->Borders->Bottom->LineStyle = BorderStyleType::DashDotDot;
        $style->Borders->Bottom->Color = Color::FromArgb(255, 0, 120, 0);
        $style->Borders->Diagonal->LineStyle = BorderStyleType::SlantedDashDot;
        $style->Borders->Diagonal->Color = Color::FromArgb(255, 128, 0, 128);
        $style->Borders->DiagonalUp = true;
        $style->Borders->DiagonalDown = true;
        $style->NumberFormat = '#,##0.00';
        $style->HorizontalAlignment = HorizontalAlignmentType::Distributed;
        $style->VerticalAlignment = VerticalAlignmentType::Distributed;
        $style->WrapText = true;
        $style->IndentLevel = 2;
        $style->TextRotation = 45;
        $style->ShrinkToFit = true;
        $style->ReadingOrder = 2;
        $style->RelativeIndent = 1;
        $style->IsLocked = false;
        $style->IsHidden = true;
    }

    public static function ApplyCustomNumberStyle($style)
    {
        $style->Font->IsBold = true;
        $style->Borders->Top->LineStyle = BorderStyleType::MediumDashed;
        $style->Borders->Top->Color = Color::FromArgb(255, 0, 0, 255);
        $style->HorizontalAlignment = HorizontalAlignmentType::CenterContinuous;
        $style->VerticalAlignment = VerticalAlignmentType::Top;
        $style->NumberFormat = '0.0000';
    }

    public static function ApplyDateStyle($style)
    {
        $style->Font->Name = 'Times New Roman';
        $style->Font->Size = 12.0;
        $style->Font->Color = Color::FromArgb(255, 0, 96, 160);
        $style->Pattern = FillPattern::Solid;
        $style->ForegroundColor = Color::FromArgb(255, 221, 235, 247);
        $style->NumberFormat = 'm/d/yyyy h:mm';
        $style->HorizontalAlignment = HorizontalAlignmentType::Center;
        $style->VerticalAlignment = VerticalAlignmentType::Center;
    }

    public static function ApplyVerticalTextStyle($style)
    {
        $style->Font->Name = 'Consolas';
        $style->Font->Size = 10.0;
        $style->Font->IsItalic = true;
        $style->Pattern = FillPattern::DarkHorizontal;
        $style->ForegroundColor = Color::FromArgb(255, 217, 217, 217);
        $style->BackgroundColor = Color::FromArgb(255, 255, 255, 255);
        $style->Borders->Left->LineStyle = BorderStyleType::Thin;
        $style->Borders->Left->Color = Color::FromArgb(255, 64, 64, 64);
        $style->HorizontalAlignment = HorizontalAlignmentType::Justify;
        $style->VerticalAlignment = VerticalAlignmentType::Justify;
        $style->WrapText = true;
        $style->TextRotation = 255;
        $style->ReadingOrder = 1;
        $style->RelativeIndent = 2;
        $style->IsLocked = true;
        $style->IsHidden = false;
    }

    public static function AssertPrimaryStyle($style)
    {
        AssertEx::Equal('Arial', $style->Font->Name);
        AssertEx::Equal(14.5, $style->Font->Size);
        AssertEx::True($style->Font->IsBold);
        AssertEx::True($style->Font->IsItalic);
        AssertEx::Equal(FontUnderlineType::Single, $style->Font->Underline);
        AssertEx::True($style->Font->IsStrikeout);
        AssertEx::Equal(Color::FromArgb(255, 17, 34, 51), $style->Font->Color);
        AssertEx::Equal(FillPattern::LightGrid, $style->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 210, 220, 30), $style->ForegroundColor);
        AssertEx::Equal(Color::FromArgb(255, 12, 45, 78), $style->BackgroundColor);
        AssertEx::Equal(BorderStyleType::Dotted, $style->Borders->Left->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 200, 0, 0), $style->Borders->Left->Color);
        AssertEx::Equal(BorderStyleType::MediumDashDot, $style->Borders->Right->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 240, 120, 0), $style->Borders->Right->Color);
        AssertEx::Equal(BorderStyleType::Double, $style->Borders->Top->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 0, 255), $style->Borders->Top->Color);
        AssertEx::Equal(BorderStyleType::DashDotDot, $style->Borders->Bottom->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 120, 0), $style->Borders->Bottom->Color);
        AssertEx::Equal(BorderStyleType::SlantedDashDot, $style->Borders->Diagonal->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 128, 0, 128), $style->Borders->Diagonal->Color);
        AssertEx::True($style->Borders->DiagonalUp);
        AssertEx::True($style->Borders->DiagonalDown);
        AssertEx::Equal(4, $style->Number);
        AssertEx::Null($style->Custom);
        AssertEx::Equal('#,##0.00', $style->NumberFormat);
        AssertEx::Equal(HorizontalAlignmentType::Distributed, $style->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Distributed, $style->VerticalAlignment);
        AssertEx::True($style->WrapText);
        AssertEx::Equal(2, $style->IndentLevel);
        AssertEx::Equal(45, $style->TextRotation);
        AssertEx::True($style->ShrinkToFit);
        AssertEx::Equal(2, $style->ReadingOrder);
        AssertEx::Equal(1, $style->RelativeIndent);
        AssertEx::False($style->IsLocked);
        AssertEx::True($style->IsHidden);
    }

    public static function AssertCustomNumberStyle($style)
    {
        AssertEx::True($style->Font->IsBold);
        AssertEx::Equal(BorderStyleType::MediumDashed, $style->Borders->Top->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 0, 255), $style->Borders->Top->Color);
        AssertEx::Equal(HorizontalAlignmentType::CenterContinuous, $style->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Top, $style->VerticalAlignment);
        AssertEx::Equal('0.0000', $style->Custom);
        AssertEx::Equal('0.0000', $style->NumberFormat);
    }

    public static function AssertDateStyle($style)
    {
        AssertEx::Equal('Times New Roman', $style->Font->Name);
        AssertEx::Equal(12.0, $style->Font->Size);
        AssertEx::Equal(Color::FromArgb(255, 0, 96, 160), $style->Font->Color);
        AssertEx::Equal(FillPattern::Solid, $style->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 221, 235, 247), $style->ForegroundColor);
        AssertEx::Equal('m/d/yyyy h:mm', $style->NumberFormat);
        AssertEx::Equal(HorizontalAlignmentType::Center, $style->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Center, $style->VerticalAlignment);
    }

    public static function AssertVerticalTextStyle($style)
    {
        AssertEx::Equal('Consolas', $style->Font->Name);
        AssertEx::Equal(10.0, $style->Font->Size);
        AssertEx::True($style->Font->IsItalic);
        AssertEx::Equal(FillPattern::DarkHorizontal, $style->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 217, 217, 217), $style->ForegroundColor);
        AssertEx::Equal(Color::FromArgb(255, 255, 255, 255), $style->BackgroundColor);
        AssertEx::Equal(BorderStyleType::Thin, $style->Borders->Left->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 64, 64, 64), $style->Borders->Left->Color);
        AssertEx::Equal(HorizontalAlignmentType::Justify, $style->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Justify, $style->VerticalAlignment);
        AssertEx::True($style->WrapText);
        AssertEx::Equal(255, $style->TextRotation);
        AssertEx::Equal(1, $style->ReadingOrder);
        AssertEx::Equal(2, $style->RelativeIndent);
        AssertEx::True($style->IsLocked);
        AssertEx::False($style->IsHidden);
    }

    public static function AssertStyleShowcase(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets['Styled'];
        self::AssertPrimaryStyle($sheet->Cells['A1']->GetStyle());
        self::AssertCustomNumberStyle($sheet->Cells['B2']->GetStyle());
        self::AssertDateStyle($sheet->Cells['C3']->GetStyle());
        self::AssertVerticalTextStyle($sheet->Cells['D4']->GetStyle());

        for ($index = 0; $index < count(self::$FillPatternShowcase); $index++) {
            $style = $sheet->Cells[[self::FillPatternRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$FillPatternShowcase[$index], $style->Pattern);
            AssertEx::Equal(self::GetShowcaseColor($index), $style->ForegroundColor);
            AssertEx::Equal(self::GetShowcaseAccentColor($index), $style->BackgroundColor);
        }

        for ($index = 0; $index < count(self::$BorderStyleShowcase); $index++) {
            $style = $sheet->Cells[[self::BorderStyleRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$BorderStyleShowcase[$index], $style->Borders->Top->LineStyle);
            AssertEx::Equal(self::GetShowcaseColor($index), $style->Borders->Top->Color);
            AssertEx::Equal(BorderStyleType::Thin, $style->Borders->Bottom->LineStyle);
            AssertEx::Equal(self::GetShowcaseAccentColor($index), $style->Borders->Bottom->Color);
        }

        for ($index = 0; $index < count(self::$HorizontalAlignmentShowcase); $index++) {
            $style = $sheet->Cells[[self::HorizontalAlignmentRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$HorizontalAlignmentShowcase[$index], $style->HorizontalAlignment);
            AssertEx::True($style->WrapText);
        }

        for ($index = 0; $index < count(self::$VerticalAlignmentShowcase); $index++) {
            $style = $sheet->Cells[[self::VerticalAlignmentRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$VerticalAlignmentShowcase[$index], $style->VerticalAlignment);
            AssertEx::True($style->ShrinkToFit);
        }

        self::AssertFontShowcase($sheet);
        self::AssertNumberFormatShowcase($sheet);
        self::AssertBorderSideShowcase($sheet);
        self::AssertRotationShowcase($sheet);
        self::AssertReadingOrderShowcase($sheet);
        self::AssertProtectionShowcase($sheet);
        self::AssertFontSizeShowcase($sheet);
        self::AssertIndentShowcase($sheet);
        self::AssertWrapShrinkShowcase($sheet);
    }

    private static function AddFillPatternShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$FillPatternShowcase); $index++) {
            $cell = $sheet->Cells[[self::FillPatternRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('P' . $index);
            $style = $cell->GetStyle();
            $style->Pattern = self::$FillPatternShowcase[$index];
            $style->ForegroundColor = self::GetShowcaseColor($index);
            $style->BackgroundColor = self::GetShowcaseAccentColor($index);
            $cell->SetStyle($style);
        }
    }

    private static function AddBorderStyleShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$BorderStyleShowcase); $index++) {
            $cell = $sheet->Cells[[self::BorderStyleRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('B' . $index);
            $style = $cell->GetStyle();
            $style->Borders->Top->LineStyle = self::$BorderStyleShowcase[$index];
            $style->Borders->Top->Color = self::GetShowcaseColor($index);
            $style->Borders->Bottom->LineStyle = BorderStyleType::Thin;
            $style->Borders->Bottom->Color = self::GetShowcaseAccentColor($index);
            $cell->SetStyle($style);
        }
    }

    private static function AddHorizontalAlignmentShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$HorizontalAlignmentShowcase); $index++) {
            $cell = $sheet->Cells[[self::HorizontalAlignmentRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('Align ' . $index . ' sample');
            $style = $cell->GetStyle();
            $style->HorizontalAlignment = self::$HorizontalAlignmentShowcase[$index];
            $style->WrapText = true;
            $style->Pattern = FillPattern::Solid;
            $style->ForegroundColor = self::GetShowcaseAccentColor($index);
            $cell->SetStyle($style);
        }
    }

    private static function AddVerticalAlignmentShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$VerticalAlignmentShowcase); $index++) {
            $cell = $sheet->Cells[[self::VerticalAlignmentRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('Vertical ' . $index);
            $style = $cell->GetStyle();
            $style->VerticalAlignment = self::$VerticalAlignmentShowcase[$index];
            $style->ShrinkToFit = true;
            $style->Pattern = FillPattern::LightTrellis;
            $style->ForegroundColor = self::GetShowcaseColor($index);
            $cell->SetStyle($style);
        }
    }

    private static function AddFontShowcase($sheet)
    {
        $nameCell = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn]];
        $nameCell->PutValue('Cambria 16');
        $nameStyle = $nameCell->GetStyle();
        $nameStyle->Font->Name = 'Cambria';
        $nameStyle->Font->Size = 16.0;
        $nameStyle->Pattern = FillPattern::Solid;
        $nameStyle->ForegroundColor = Color::FromArgb(255, 242, 242, 242);
        $nameCell->SetStyle($nameStyle);

        $boldCell = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 1]];
        $boldCell->PutValue('Bold');
        $boldStyle = $boldCell->GetStyle();
        $boldStyle->Font->IsBold = true;
        $boldStyle->Font->Size = 12.0;
        $boldStyle->ForegroundColor = Color::FromArgb(255, 31, 78, 121);
        $boldCell->SetStyle($boldStyle);

        $italicCell = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 2]];
        $italicCell->PutValue('Italic');
        $italicStyle = $italicCell->GetStyle();
        $italicStyle->Font->IsItalic = true;
        $italicStyle->Font->Size = 12.0;
        $italicStyle->ForegroundColor = Color::FromArgb(255, 128, 100, 162);
        $italicCell->SetStyle($italicStyle);

        $underlineCell = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 3]];
        $underlineCell->PutValue('Underline');
        $underlineStyle = $underlineCell->GetStyle();
        $underlineStyle->Font->Underline = FontUnderlineType::Single;
        $underlineStyle->Font->Color = Color::FromArgb(255, 0, 112, 192);
        $underlineCell->SetStyle($underlineStyle);

        $strikeCell = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 4]];
        $strikeCell->PutValue('Strike');
        $strikeStyle = $strikeCell->GetStyle();
        $strikeStyle->Font->IsStrikeout = true;
        $strikeStyle->Font->Color = Color::FromArgb(255, 192, 0, 0);
        $strikeCell->SetStyle($strikeStyle);

        $colorCell = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 5]];
        $colorCell->PutValue('Color');
        $colorStyle = $colorCell->GetStyle();
        $colorStyle->Font->IsBold = true;
        $colorStyle->Font->Color = Color::FromArgb(255, 0, 176, 80);
        $colorStyle->Pattern = FillPattern::Solid;
        $colorStyle->ForegroundColor = Color::FromArgb(255, 226, 239, 218);
        $colorCell->SetStyle($colorStyle);
    }

    private static function AddNumberFormatShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$NumberFormatShowcase); $index++) {
            $cell = $sheet->Cells[[self::NumberFormatRow, self::ShowcaseStartColumn + $index]];
            if ($index === 3) {
                $cell->PutValue(0.375);
            } elseif ($index === 5) {
                $cell->PutValue(12.75);
            } else {
                $cell->PutValue(1234.567 + $index);
            }

            $style = $cell->GetStyle();
            $style->NumberFormat = self::$NumberFormatShowcase[$index];
            $style->HorizontalAlignment = HorizontalAlignmentType::Right;
            $style->Pattern = FillPattern::Solid;
            $style->ForegroundColor = self::GetShowcaseAccentColor($index + 10);
            $cell->SetStyle($style);
        }
    }

    private static function AddBorderSideShowcase($sheet)
    {
        $leftCell = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn]];
        $leftCell->PutValue('Left');
        $leftStyle = $leftCell->GetStyle();
        $leftStyle->Borders->Left->LineStyle = BorderStyleType::Thick;
        $leftStyle->Borders->Left->Color = Color::FromArgb(255, 192, 0, 0);
        $leftStyle->Pattern = FillPattern::Solid;
        $leftStyle->ForegroundColor = Color::FromArgb(255, 255, 242, 204);
        $leftCell->SetStyle($leftStyle);

        $rightCell = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 1]];
        $rightCell->PutValue('Right');
        $rightStyle = $rightCell->GetStyle();
        $rightStyle->Borders->Right->LineStyle = BorderStyleType::Double;
        $rightStyle->Borders->Right->Color = Color::FromArgb(255, 0, 112, 192);
        $rightStyle->Pattern = FillPattern::Solid;
        $rightStyle->ForegroundColor = Color::FromArgb(255, 221, 235, 247);
        $rightCell->SetStyle($rightStyle);

        $topCell = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 2]];
        $topCell->PutValue('Top');
        $topStyle = $topCell->GetStyle();
        $topStyle->Borders->Top->LineStyle = BorderStyleType::DashDot;
        $topStyle->Borders->Top->Color = Color::FromArgb(255, 112, 48, 160);
        $topStyle->Pattern = FillPattern::Solid;
        $topStyle->ForegroundColor = Color::FromArgb(255, 234, 209, 220);
        $topCell->SetStyle($topStyle);

        $bottomCell = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 3]];
        $bottomCell->PutValue('Bottom');
        $bottomStyle = $bottomCell->GetStyle();
        $bottomStyle->Borders->Bottom->LineStyle = BorderStyleType::MediumDashed;
        $bottomStyle->Borders->Bottom->Color = Color::FromArgb(255, 0, 176, 80);
        $bottomStyle->Pattern = FillPattern::Solid;
        $bottomStyle->ForegroundColor = Color::FromArgb(255, 226, 239, 218);
        $bottomCell->SetStyle($bottomStyle);

        $diagonalUpCell = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 4]];
        $diagonalUpCell->PutValue('Diag Up');
        $diagonalUpStyle = $diagonalUpCell->GetStyle();
        $diagonalUpStyle->Borders->Diagonal->LineStyle = BorderStyleType::SlantedDashDot;
        $diagonalUpStyle->Borders->Diagonal->Color = Color::FromArgb(255, 255, 0, 0);
        $diagonalUpStyle->Borders->DiagonalUp = true;
        $diagonalUpStyle->Pattern = FillPattern::Solid;
        $diagonalUpStyle->ForegroundColor = Color::FromArgb(255, 252, 228, 214);
        $diagonalUpCell->SetStyle($diagonalUpStyle);

        $diagonalDownCell = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 5]];
        $diagonalDownCell->PutValue('Diag Down');
        $diagonalDownStyle = $diagonalDownCell->GetStyle();
        $diagonalDownStyle->Borders->Diagonal->LineStyle = BorderStyleType::SlantedDashDot;
        $diagonalDownStyle->Borders->Diagonal->Color = Color::FromArgb(255, 0, 176, 240);
        $diagonalDownStyle->Borders->DiagonalDown = true;
        $diagonalDownStyle->Pattern = FillPattern::Solid;
        $diagonalDownStyle->ForegroundColor = Color::FromArgb(255, 222, 234, 246);
        $diagonalDownCell->SetStyle($diagonalDownStyle);
    }

    private static function AddRotationShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$RotationShowcase); $index++) {
            $cell = $sheet->Cells[[self::RotationRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('Rot ' . self::$RotationShowcase[$index]);
            $style = $cell->GetStyle();
            $style->TextRotation = self::$RotationShowcase[$index];
            $style->HorizontalAlignment = HorizontalAlignmentType::Center;
            $style->VerticalAlignment = VerticalAlignmentType::Center;
            $style->WrapText = true;
            $style->Pattern = FillPattern::Solid;
            $style->ForegroundColor = self::GetShowcaseColor($index + 20);
            $cell->SetStyle($style);
        }
    }

    private static function AddReadingOrderShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$ReadingOrderShowcase); $index++) {
            $cell = $sheet->Cells[[self::ReadingOrderRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('Order ' . self::$ReadingOrderShowcase[$index]);
            $style = $cell->GetStyle();
            $style->ReadingOrder = self::$ReadingOrderShowcase[$index];
            $style->IndentLevel = $index;
            if (self::$ReadingOrderShowcase[$index] > 0) {
                $style->RelativeIndent = $index + 1;
            }
            $style->HorizontalAlignment = HorizontalAlignmentType::Distributed;
            $style->Pattern = FillPattern::Solid;
            $style->ForegroundColor = self::GetShowcaseAccentColor($index + 20);
            $cell->SetStyle($style);
        }
    }

    private static function AddProtectionShowcase($sheet)
    {
        $lockedCell = $sheet->Cells[[self::ProtectionRow, self::ShowcaseStartColumn]];
        $lockedCell->PutValue('Locked');
        $lockedStyle = $lockedCell->GetStyle();
        $lockedStyle->IsLocked = true;
        $lockedStyle->IsHidden = false;
        $lockedStyle->Pattern = FillPattern::Solid;
        $lockedStyle->ForegroundColor = Color::FromArgb(255, 226, 239, 218);
        $lockedCell->SetStyle($lockedStyle);

        $unlockedCell = $sheet->Cells[[self::ProtectionRow, self::ShowcaseStartColumn + 1]];
        $unlockedCell->PutValue('Unlocked');
        $unlockedStyle = $unlockedCell->GetStyle();
        $unlockedStyle->IsLocked = false;
        $unlockedStyle->IsHidden = false;
        $unlockedStyle->Pattern = FillPattern::Solid;
        $unlockedStyle->ForegroundColor = Color::FromArgb(255, 255, 242, 204);
        $unlockedCell->SetStyle($unlockedStyle);

        $hiddenCell = $sheet->Cells[[self::ProtectionRow, self::ShowcaseStartColumn + 2]];
        $hiddenCell->PutValue('Hidden');
        $hiddenStyle = $hiddenCell->GetStyle();
        $hiddenStyle->IsLocked = true;
        $hiddenStyle->IsHidden = true;
        $hiddenStyle->Pattern = FillPattern::Solid;
        $hiddenStyle->ForegroundColor = Color::FromArgb(255, 242, 220, 219);
        $hiddenCell->SetStyle($hiddenStyle);
    }

    private static function AddFontSizeShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$FontSizeShowcase); $index++) {
            $cell = $sheet->Cells[[self::FontSizeRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('Size ' . number_format(self::$FontSizeShowcase[$index], 0, '.', ''));
            $style = $cell->GetStyle();
            $style->Font->Name = 'Calibri';
            $style->Font->Size = self::$FontSizeShowcase[$index];
            $style->Pattern = FillPattern::Solid;
            $style->ForegroundColor = self::GetShowcaseAccentColor($index + 30);
            $cell->SetStyle($style);
        }
    }

    private static function AddIndentShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$IndentLevelShowcase); $index++) {
            $cell = $sheet->Cells[[self::IndentRow, self::ShowcaseStartColumn + $index]];
            $cell->PutValue('Indent ' . self::$IndentLevelShowcase[$index]);
            $style = $cell->GetStyle();
            $style->HorizontalAlignment = HorizontalAlignmentType::Left;
            $style->IndentLevel = self::$IndentLevelShowcase[$index];
            $style->RelativeIndent = self::$IndentLevelShowcase[$index];
            $style->Pattern = FillPattern::Solid;
            $style->ForegroundColor = self::GetShowcaseColor($index + 30);
            $cell->SetStyle($style);
        }
    }

    private static function AddWrapShrinkShowcase($sheet)
    {
        $wrapCell = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn]];
        $wrapCell->PutValue('Wrap text sample');
        $wrapStyle = $wrapCell->GetStyle();
        $wrapStyle->WrapText = true;
        $wrapStyle->HorizontalAlignment = HorizontalAlignmentType::Justify;
        $wrapStyle->Pattern = FillPattern::Solid;
        $wrapStyle->ForegroundColor = Color::FromArgb(255, 252, 228, 214);
        $wrapCell->SetStyle($wrapStyle);

        $shrinkCell = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn + 1]];
        $shrinkCell->PutValue('Shrink to fit sample');
        $shrinkStyle = $shrinkCell->GetStyle();
        $shrinkStyle->ShrinkToFit = true;
        $shrinkStyle->HorizontalAlignment = HorizontalAlignmentType::Center;
        $shrinkStyle->Pattern = FillPattern::Solid;
        $shrinkStyle->ForegroundColor = Color::FromArgb(255, 221, 235, 247);
        $shrinkCell->SetStyle($shrinkStyle);

        $wrapShrinkCell = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn + 2]];
        $wrapShrinkCell->PutValue('Wrap and shrink');
        $wrapShrinkStyle = $wrapShrinkCell->GetStyle();
        $wrapShrinkStyle->WrapText = true;
        $wrapShrinkStyle->ShrinkToFit = true;
        $wrapShrinkStyle->HorizontalAlignment = HorizontalAlignmentType::Distributed;
        $wrapShrinkStyle->VerticalAlignment = VerticalAlignmentType::Center;
        $wrapShrinkStyle->Pattern = FillPattern::Solid;
        $wrapShrinkStyle->ForegroundColor = Color::FromArgb(255, 226, 239, 218);
        $wrapShrinkCell->SetStyle($wrapShrinkStyle);

        $justifyCell = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn + 3]];
        $justifyCell->PutValue('Distributed indent');
        $justifyStyle = $justifyCell->GetStyle();
        $justifyStyle->WrapText = true;
        $justifyStyle->HorizontalAlignment = HorizontalAlignmentType::Distributed;
        $justifyStyle->IndentLevel = 3;
        $justifyStyle->RelativeIndent = 2;
        $justifyStyle->Pattern = FillPattern::Solid;
        $justifyStyle->ForegroundColor = Color::FromArgb(255, 242, 242, 242);
        $justifyCell->SetStyle($justifyStyle);
    }

    private static function AssertFontShowcase($sheet)
    {
        $nameStyle = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn]]->GetStyle();
        AssertEx::Equal('Cambria', $nameStyle->Font->Name);
        AssertEx::Equal(16.0, $nameStyle->Font->Size);
        AssertEx::Equal(FillPattern::Solid, $nameStyle->Pattern);

        $boldStyle = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 1]]->GetStyle();
        AssertEx::True($boldStyle->Font->IsBold);
        AssertEx::Equal(12.0, $boldStyle->Font->Size);
        AssertEx::Equal(Color::FromArgb(255, 31, 78, 121), $boldStyle->ForegroundColor);

        $italicStyle = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 2]]->GetStyle();
        AssertEx::True($italicStyle->Font->IsItalic);
        AssertEx::Equal(12.0, $italicStyle->Font->Size);
        AssertEx::Equal(Color::FromArgb(255, 128, 100, 162), $italicStyle->ForegroundColor);

        $underlineStyle = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 3]]->GetStyle();
        AssertEx::Equal(FontUnderlineType::Single, $underlineStyle->Font->Underline);
        AssertEx::Equal(Color::FromArgb(255, 0, 112, 192), $underlineStyle->Font->Color);

        $strikeStyle = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 4]]->GetStyle();
        AssertEx::True($strikeStyle->Font->IsStrikeout);
        AssertEx::Equal(Color::FromArgb(255, 192, 0, 0), $strikeStyle->Font->Color);

        $colorStyle = $sheet->Cells[[self::FontRow, self::ShowcaseStartColumn + 5]]->GetStyle();
        AssertEx::True($colorStyle->Font->IsBold);
        AssertEx::Equal(Color::FromArgb(255, 0, 176, 80), $colorStyle->Font->Color);
        AssertEx::Equal(FillPattern::Solid, $colorStyle->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 226, 239, 218), $colorStyle->ForegroundColor);
    }

    private static function AssertNumberFormatShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$NumberFormatShowcase); $index++) {
            $style = $sheet->Cells[[self::NumberFormatRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$NumberFormatShowcase[$index], $style->NumberFormat);
            AssertEx::Equal(HorizontalAlignmentType::Right, $style->HorizontalAlignment);
            AssertEx::Equal(FillPattern::Solid, $style->Pattern);
            AssertEx::Equal(self::GetShowcaseAccentColor($index + 10), $style->ForegroundColor);
        }
    }

    private static function AssertBorderSideShowcase($sheet)
    {
        $leftStyle = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn]]->GetStyle();
        AssertEx::Equal(BorderStyleType::Thick, $leftStyle->Borders->Left->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 192, 0, 0), $leftStyle->Borders->Left->Color);

        $rightStyle = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 1]]->GetStyle();
        AssertEx::Equal(BorderStyleType::Double, $rightStyle->Borders->Right->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 112, 192), $rightStyle->Borders->Right->Color);

        $topStyle = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 2]]->GetStyle();
        AssertEx::Equal(BorderStyleType::DashDot, $topStyle->Borders->Top->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 112, 48, 160), $topStyle->Borders->Top->Color);

        $bottomStyle = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 3]]->GetStyle();
        AssertEx::Equal(BorderStyleType::MediumDashed, $bottomStyle->Borders->Bottom->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 176, 80), $bottomStyle->Borders->Bottom->Color);

        $diagonalUpStyle = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 4]]->GetStyle();
        AssertEx::Equal(BorderStyleType::SlantedDashDot, $diagonalUpStyle->Borders->Diagonal->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 255, 0, 0), $diagonalUpStyle->Borders->Diagonal->Color);
        AssertEx::True($diagonalUpStyle->Borders->DiagonalUp);
        AssertEx::False($diagonalUpStyle->Borders->DiagonalDown);

        $diagonalDownStyle = $sheet->Cells[[self::BorderSidesRow, self::ShowcaseStartColumn + 5]]->GetStyle();
        AssertEx::Equal(BorderStyleType::SlantedDashDot, $diagonalDownStyle->Borders->Diagonal->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 176, 240), $diagonalDownStyle->Borders->Diagonal->Color);
        AssertEx::False($diagonalDownStyle->Borders->DiagonalUp);
        AssertEx::True($diagonalDownStyle->Borders->DiagonalDown);
    }

    private static function AssertRotationShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$RotationShowcase); $index++) {
            $style = $sheet->Cells[[self::RotationRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$RotationShowcase[$index], $style->TextRotation);
            AssertEx::Equal(HorizontalAlignmentType::Center, $style->HorizontalAlignment);
            AssertEx::Equal(VerticalAlignmentType::Center, $style->VerticalAlignment);
            AssertEx::True($style->WrapText);
        }
    }

    private static function AssertReadingOrderShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$ReadingOrderShowcase); $index++) {
            $style = $sheet->Cells[[self::ReadingOrderRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(self::$ReadingOrderShowcase[$index], $style->ReadingOrder);
            AssertEx::Equal($index, $style->IndentLevel);
            AssertEx::Equal(self::$ReadingOrderShowcase[$index] > 0 ? $index + 1 : 0, $style->RelativeIndent);
            AssertEx::Equal(HorizontalAlignmentType::Distributed, $style->HorizontalAlignment);
            AssertEx::Equal(FillPattern::Solid, $style->Pattern);
            AssertEx::Equal(self::GetShowcaseAccentColor($index + 20), $style->ForegroundColor);
        }
    }

    private static function AssertProtectionShowcase($sheet)
    {
        $lockedStyle = $sheet->Cells[[self::ProtectionRow, self::ShowcaseStartColumn]]->GetStyle();
        AssertEx::True($lockedStyle->IsLocked);
        AssertEx::False($lockedStyle->IsHidden);

        $unlockedStyle = $sheet->Cells[[self::ProtectionRow, self::ShowcaseStartColumn + 1]]->GetStyle();
        AssertEx::False($unlockedStyle->IsLocked);
        AssertEx::False($unlockedStyle->IsHidden);

        $hiddenStyle = $sheet->Cells[[self::ProtectionRow, self::ShowcaseStartColumn + 2]]->GetStyle();
        AssertEx::True($hiddenStyle->IsLocked);
        AssertEx::True($hiddenStyle->IsHidden);
    }

    private static function AssertFontSizeShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$FontSizeShowcase); $index++) {
            $style = $sheet->Cells[[self::FontSizeRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal('Calibri', $style->Font->Name);
            AssertEx::Equal(self::$FontSizeShowcase[$index], $style->Font->Size);
            AssertEx::Equal(FillPattern::Solid, $style->Pattern);
            AssertEx::Equal(self::GetShowcaseAccentColor($index + 30), $style->ForegroundColor);
        }
    }

    private static function AssertIndentShowcase($sheet)
    {
        for ($index = 0; $index < count(self::$IndentLevelShowcase); $index++) {
            $style = $sheet->Cells[[self::IndentRow, self::ShowcaseStartColumn + $index]]->GetStyle();
            AssertEx::Equal(HorizontalAlignmentType::Left, $style->HorizontalAlignment);
            AssertEx::Equal(self::$IndentLevelShowcase[$index], $style->IndentLevel);
            AssertEx::Equal(FillPattern::Solid, $style->Pattern);
            AssertEx::Equal(self::GetShowcaseColor($index + 30), $style->ForegroundColor);
        }
    }

    private static function AssertWrapShrinkShowcase($sheet)
    {
        $wrapStyle = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn]]->GetStyle();
        AssertEx::True($wrapStyle->WrapText);
        AssertEx::Equal(HorizontalAlignmentType::Justify, $wrapStyle->HorizontalAlignment);

        $shrinkStyle = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn + 1]]->GetStyle();
        AssertEx::True($shrinkStyle->ShrinkToFit);
        AssertEx::Equal(HorizontalAlignmentType::Center, $shrinkStyle->HorizontalAlignment);

        $wrapShrinkStyle = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn + 2]]->GetStyle();
        AssertEx::True($wrapShrinkStyle->WrapText);
        AssertEx::True($wrapShrinkStyle->ShrinkToFit);
        AssertEx::Equal(HorizontalAlignmentType::Distributed, $wrapShrinkStyle->HorizontalAlignment);
        AssertEx::Equal(VerticalAlignmentType::Center, $wrapShrinkStyle->VerticalAlignment);

        $justifyStyle = $sheet->Cells[[self::WrapShrinkRow, self::ShowcaseStartColumn + 3]]->GetStyle();
        AssertEx::True($justifyStyle->WrapText);
        AssertEx::Equal(HorizontalAlignmentType::Distributed, $justifyStyle->HorizontalAlignment);
        AssertEx::Equal(3, $justifyStyle->IndentLevel);
        AssertEx::Equal(2, $justifyStyle->RelativeIndent);
    }

    private static function AddShowcaseLabel($sheet, $rowIndex, $label)
    {
        $cell = $sheet->Cells[[$rowIndex, self::ShowcaseStartColumn - 1]];
        $cell->PutValue($label);
        $style = $cell->GetStyle();
        $style->Font->IsBold = true;
        $style->Pattern = FillPattern::Solid;
        $style->ForegroundColor = Color::FromArgb(255, 217, 225, 242);
        $style->Borders->Bottom->LineStyle = BorderStyleType::Thin;
        $style->Borders->Bottom->Color = Color::FromArgb(255, 79, 129, 189);
        $cell->SetStyle($style);
    }

    private static function GetShowcaseColor($index)
    {
        return Color::FromArgb(
            255,
            (40 + ($index * 29)) % 256,
            (90 + ($index * 37)) % 256,
            (140 + ($index * 43)) % 256
        );
    }

    private static function GetShowcaseAccentColor($index)
    {
        return Color::FromArgb(
            255,
            (220 + ($index * 17)) % 256,
            (180 + ($index * 19)) % 256,
            (120 + ($index * 23)) % 256
        );
    }
}

final class ValidationScenarioFactory
{
    public static function CreateValidationWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Validation Sheet';

        $sheet->Cells['A1']->PutValue('Open');
        $sheet->Cells['B2']->PutValue(5);
        $sheet->Cells['C3']->PutValue(7);
        $sheet->Cells['E2']->PutValue(8);
        $sheet->Cells['G1']->PutValue('ABCDE');

        $listIndex = $sheet->Validations->Add(CellArea::CreateCellArea('A1', 'A3'));
        $listValidation = $sheet->Validations[$listIndex];
        $listValidation->Type = ValidationType::List;
        $listValidation->Formula1 = '"Open,Closed"';
        $listValidation->IgnoreBlank = true;
        $listValidation->InCellDropDown = true;
        $listValidation->ShowInput = true;
        $listValidation->InputTitle = 'Status';
        $listValidation->InputMessage = 'Pick a status';
        $listValidation->ShowError = true;
        $listValidation->ErrorTitle = 'Invalid';
        $listValidation->ErrorMessage = 'Choose from the list';

        $decimalIndex = $sheet->Validations->Add(CellArea::CreateCellArea('B2', 'C3'));
        $decimalValidation = $sheet->Validations[$decimalIndex];
        $decimalValidation->Type = ValidationType::Decimal;
        $decimalValidation->Operator = OperatorType::Between;
        $decimalValidation->Formula1 = '1.5';
        $decimalValidation->Formula2 = '9.5';
        $decimalValidation->IgnoreBlank = false;
        $decimalValidation->InCellDropDown = false;
        $decimalValidation->ShowError = true;
        $decimalValidation->ErrorTitle = 'Range';
        $decimalValidation->ErrorMessage = 'Enter 1.5-9.5';
        $decimalValidation->AddArea(CellArea::CreateCellArea('E2', 'E3'));

        $customIndex = $sheet->Validations->Add(CellArea::CreateCellArea('G1', 'G1'));
        $customValidation = $sheet->Validations[$customIndex];
        $customValidation->Type = ValidationType::Custom;
        $customValidation->AlertStyle = ValidationAlertType::Warning;
        $customValidation->Formula1 = 'LEN(G1)<=5';
        $customValidation->ShowInput = true;
        $customValidation->InputTitle = 'Code';
        $customValidation->InputMessage = 'Up to 5 chars';

        return $workbook;
    }

    public static function AssertValidations(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets['Validation Sheet'];
        AssertEx::Equal(3, $sheet->Validations->Count);

        $listValidation = self::GetValidationByArea($sheet->Validations, 'A1:A3');
        AssertEx::Equal(ValidationType::List, $listValidation->Type);
        AssertEx::Equal(OperatorType::None, $listValidation->Operator);
        AssertEx::Equal(1, count($listValidation->Areas));
        self::AssertArea($listValidation->Areas[0], 0, 0, 3, 1);
        AssertEx::Equal('"Open,Closed"', $listValidation->Formula1);
        AssertEx::Equal('', $listValidation->Formula2);
        AssertEx::True($listValidation->IgnoreBlank);
        AssertEx::True($listValidation->InCellDropDown);
        AssertEx::True($listValidation->ShowInput);
        AssertEx::True($listValidation->ShowError);
        AssertEx::Equal('Status', $listValidation->InputTitle);
        AssertEx::Equal('Pick a status', $listValidation->InputMessage);
        AssertEx::Equal('Invalid', $listValidation->ErrorTitle);
        AssertEx::Equal('Choose from the list', $listValidation->ErrorMessage);

        $decimalValidation = self::GetValidationByArea($sheet->Validations, 'B2:C3 E2:E3');
        AssertEx::Equal(ValidationType::Decimal, $decimalValidation->Type);
        AssertEx::Equal(OperatorType::Between, $decimalValidation->Operator);
        AssertEx::Equal(2, count($decimalValidation->Areas));
        self::AssertArea($decimalValidation->Areas[0], 1, 1, 2, 2);
        self::AssertArea($decimalValidation->Areas[1], 1, 4, 2, 1);
        AssertEx::Equal('1.5', $decimalValidation->Formula1);
        AssertEx::Equal('9.5', $decimalValidation->Formula2);
        AssertEx::False($decimalValidation->IgnoreBlank);
        AssertEx::False($decimalValidation->InCellDropDown);
        AssertEx::False($decimalValidation->ShowInput);
        AssertEx::True($decimalValidation->ShowError);
        AssertEx::Equal('Range', $decimalValidation->ErrorTitle);
        AssertEx::Equal('Enter 1.5-9.5', $decimalValidation->ErrorMessage);

        $customValidation = self::GetValidationByArea($sheet->Validations, 'G1');
        AssertEx::Equal(ValidationType::Custom, $customValidation->Type);
        AssertEx::Equal(ValidationAlertType::Warning, $customValidation->AlertStyle);
        AssertEx::Equal(OperatorType::None, $customValidation->Operator);
        AssertEx::Equal(1, count($customValidation->Areas));
        self::AssertArea($customValidation->Areas[0], 0, 6, 1, 1);
        AssertEx::Equal('LEN(G1)<=5', $customValidation->Formula1);
        AssertEx::Equal('', $customValidation->Formula2);
        AssertEx::True($customValidation->ShowInput);
        AssertEx::False($customValidation->ShowError);
        AssertEx::Equal('Code', $customValidation->InputTitle);
        AssertEx::Equal('Up to 5 chars', $customValidation->InputMessage);

        $a1Validation = $sheet->Validations->GetValidationInCell(0, 0);
        AssertEx::NotNull($a1Validation);
        AssertEx::Equal(ValidationType::List, $a1Validation->Type);

        $c2Validation = $sheet->Validations->GetValidationInCell(1, 2);
        AssertEx::NotNull($c2Validation);
        AssertEx::Equal(ValidationType::Decimal, $c2Validation->Type);

        $missingValidation = $sheet->Validations->GetValidationInCell(9, 9);
        AssertEx::Null($missingValidation);
    }

    private static function GetValidationByArea($validations, $expectedArea)
    {
        for ($index = 0; $index < $validations->Count; $index++) {
            $validation = $validations[$index];
            if (self::BuildAreaKey($validation) === $expectedArea) {
                return $validation;
            }
        }

        throw new \RuntimeException("Validation with area '" . $expectedArea . "' was not found.");
    }

    private static function BuildAreaKey($validation)
    {
        $areas = array();
        for ($index = 0; $index < count($validation->Areas); $index++) {
            $area = $validation->Areas[$index];
            $startCell = self::BuildCellName($area->StartRow, $area->StartColumn);
            if ($area->EndRow === $area->StartRow && $area->EndColumn === $area->StartColumn) {
                $areas[] = $startCell;
                continue;
            }

            $endCell = self::BuildCellName($area->EndRow, $area->EndColumn);
            $areas[] = $startCell . ':' . $endCell;
        }

        return implode(' ', $areas);
    }

    private static function BuildCellName($rowIndex, $columnIndex)
    {
        $dividend = $columnIndex + 1;
        $name = '';
        while ($dividend > 0) {
            $remainder = ($dividend - 1) % 26;
            $name = chr(ord('A') + $remainder) . $name;
            $dividend = intdiv($dividend - $remainder - 1, 26);
        }

        return $name . ($rowIndex + 1);
    }

    private static function AssertArea($area, $firstRow, $firstColumn, $totalRows, $totalColumns)
    {
        AssertEx::Equal($firstRow, $area->StartRow);
        AssertEx::Equal($firstColumn, $area->StartColumn);
        AssertEx::Equal($firstRow + $totalRows - 1, $area->EndRow);
        AssertEx::Equal($firstColumn + $totalColumns - 1, $area->EndColumn);
    }
}

final class ConditionalFormattingScenarioFactory
{
    public static function CreateConditionalFormattingWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Conditional Formatting';

        $sheet->Cells['A1']->PutValue(1);
        $sheet->Cells['A2']->PutValue(5);
        $sheet->Cells['A3']->PutValue(10);
        $sheet->Cells['A4']->PutValue(15);
        $sheet->Cells['A5']->PutValue(20);
        $sheet->Cells['C1']->PutValue(11);
        $sheet->Cells['C2']->PutValue(8);
        $sheet->Cells['C3']->PutValue(15);

        $primaryIndex = $sheet->ConditionalFormattings->Add();
        $primary = $sheet->ConditionalFormattings[$primaryIndex];
        $primary->AddArea(CellArea::CreateCellArea('A1', 'A5'));
        $betweenIndex = $primary->AddCondition(FormatConditionType::CellValue, OperatorType::Between, '5', '15');
        $between = $primary[$betweenIndex];
        $between->StopIfTrue = true;
        $between->Priority = 1;
        $betweenStyle = $between->Style;
        $betweenStyle->Pattern = FillPattern::Solid;
        $betweenStyle->ForegroundColor = Color::FromArgb(255, 255, 199, 206);
        $betweenStyle->Font->IsBold = true;
        $betweenStyle->Font->Color = Color::FromArgb(255, 156, 0, 6);
        $between->Style = $betweenStyle;

        $expressionIndex = $primary->AddCondition(FormatConditionType::Expression, OperatorType::None, 'MOD(A1,2)=0', '');
        $expression = $primary[$expressionIndex];
        $expression->Priority = 2;
        $expressionStyle = $expression->Style;
        $expressionStyle->Font->IsItalic = true;
        $expressionStyle->Font->Color = Color::FromArgb(255, 0, 0, 255);
        $expressionStyle->Borders->Bottom->LineStyle = BorderStyleType::Thin;
        $expressionStyle->Borders->Bottom->Color = Color::FromArgb(255, 0, 0, 255);
        $expression->Style = $expressionStyle;

        $secondaryIndex = $sheet->ConditionalFormattings->Add();
        $secondary = $sheet->ConditionalFormattings[$secondaryIndex];
        $secondary->AddArea(CellArea::CreateCellArea('C1', 'C3'));
        $greaterThanIndex = $secondary->AddCondition(FormatConditionType::CellValue, OperatorType::GreaterThan, '10', '');
        $greaterThan = $secondary[$greaterThanIndex];
        $greaterThan->Priority = 3;
        $greaterThanStyle = $greaterThan->Style;
        $greaterThanStyle->Pattern = FillPattern::Solid;
        $greaterThanStyle->ForegroundColor = Color::FromArgb(255, 198, 239, 206);
        $greaterThanStyle->Font->Color = Color::FromArgb(255, 0, 97, 0);
        $greaterThan->Style = $greaterThanStyle;

        return $workbook;
    }

    public static function AssertConditionalFormattings(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets['Conditional Formatting'];
        AssertEx::Equal(2, $sheet->ConditionalFormattings->Count);

        $primary = self::GetCollectionByArea($sheet->ConditionalFormattings, 'A1:A5');
        AssertEx::Equal(2, $primary->Count);
        AssertEx::Equal(1, $primary->RangeCount);
        self::AssertArea($primary->GetCellArea(0), 0, 0, 5, 1);

        $between = $primary[0];
        AssertEx::Equal(FormatConditionType::CellValue, $between->Type);
        AssertEx::Equal(OperatorType::Between, $between->Operator);
        AssertEx::Equal('5', $between->Formula1);
        AssertEx::Equal('15', $between->Formula2);
        AssertEx::Equal(1, $between->Priority);
        AssertEx::True($between->StopIfTrue);
        AssertEx::Equal(FillPattern::Solid, $between->Style->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 255, 199, 206), $between->Style->ForegroundColor);
        AssertEx::True($between->Style->Font->IsBold);
        AssertEx::Equal(Color::FromArgb(255, 156, 0, 6), $between->Style->Font->Color);

        $expression = $primary[1];
        AssertEx::Equal(FormatConditionType::Expression, $expression->Type);
        AssertEx::Equal(OperatorType::None, $expression->Operator);
        AssertEx::Equal('MOD(A1,2)=0', $expression->Formula1);
        AssertEx::Equal('', $expression->Formula2);
        AssertEx::Equal(2, $expression->Priority);
        AssertEx::False($expression->StopIfTrue);
        AssertEx::True($expression->Style->Font->IsItalic);
        AssertEx::Equal(Color::FromArgb(255, 0, 0, 255), $expression->Style->Font->Color);
        AssertEx::Equal(BorderStyleType::Thin, $expression->Style->Borders->Bottom->LineStyle);
        AssertEx::Equal(Color::FromArgb(255, 0, 0, 255), $expression->Style->Borders->Bottom->Color);

        $secondary = self::GetCollectionByArea($sheet->ConditionalFormattings, 'C1:C3');
        AssertEx::Equal(1, $secondary->Count);
        AssertEx::Equal(1, $secondary->RangeCount);
        self::AssertArea($secondary->GetCellArea(0), 0, 2, 3, 1);
        $greaterThan = $secondary[0];
        AssertEx::Equal(FormatConditionType::CellValue, $greaterThan->Type);
        AssertEx::Equal(OperatorType::GreaterThan, $greaterThan->Operator);
        AssertEx::Equal('10', $greaterThan->Formula1);
        AssertEx::Equal('', $greaterThan->Formula2);
        AssertEx::Equal(3, $greaterThan->Priority);
        AssertEx::Equal(FillPattern::Solid, $greaterThan->Style->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 198, 239, 206), $greaterThan->Style->ForegroundColor);
        AssertEx::Equal(Color::FromArgb(255, 0, 97, 0), $greaterThan->Style->Font->Color);
    }

    private static function GetCollectionByArea($collections, $expectedArea)
    {
        for ($index = 0; $index < $collections->Count; $index++) {
            $collection = $collections[$index];
            if (self::BuildAreaKey($collection) === $expectedArea) {
                return $collection;
            }
        }

        throw new \RuntimeException("Conditional formatting collection with area '" . $expectedArea . "' was not found.");
    }

    private static function BuildAreaKey($collection)
    {
        $areas = array();
        for ($index = 0; $index < $collection->RangeCount; $index++) {
            $area = $collection->GetCellArea($index);
            $startCell = self::BuildCellName($area->StartRow, $area->StartColumn);
            if ($area->EndRow === $area->StartRow && $area->EndColumn === $area->StartColumn) {
                $areas[] = $startCell;
                continue;
            }

            $endCell = self::BuildCellName($area->EndRow, $area->EndColumn);
            $areas[] = $startCell . ':' . $endCell;
        }

        return implode(' ', $areas);
    }

    private static function BuildCellName($rowIndex, $columnIndex)
    {
        $dividend = $columnIndex + 1;
        $name = '';
        while ($dividend > 0) {
            $remainder = ($dividend - 1) % 26;
            $name = chr(ord('A') + $remainder) . $name;
            $dividend = ($dividend - $remainder - 1) / 26;
        }

        return $name . ($rowIndex + 1);
    }

    private static function AssertArea($area, $firstRow, $firstColumn, $totalRows, $totalColumns)
    {
        AssertEx::Equal($firstRow, $area->StartRow);
        AssertEx::Equal($firstColumn, $area->StartColumn);
        AssertEx::Equal($firstRow + $totalRows - 1, $area->EndRow);
        AssertEx::Equal($firstColumn + $totalColumns - 1, $area->EndColumn);
    }

    public static function CreateAdvancedConditionalFormattingWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Advanced Conditional Formatting';

        $sheet->Cells['A1']->PutValue('error one');
        $sheet->Cells['A2']->PutValue('ok');
        $sheet->Cells['B1']->PutValue('ok');
        $sheet->Cells['B2']->PutValue('warning');
        $sheet->Cells['C1']->PutValue('prefix-item');
        $sheet->Cells['D1']->PutValue('item-suffix');
        $sheet->Cells['E1']->PutValue(new \DateTime('today'));
        $sheet->Cells['F1']->PutValue(1);
        $sheet->Cells['F2']->PutValue(1);
        $sheet->Cells['G1']->PutValue(10);
        $sheet->Cells['G2']->PutValue(20);
        for ($index = 1; $index <= 10; $index++) {
            $sheet->Cells['H' . $index]->PutValue($index * 10);
            $sheet->Cells['I' . $index]->PutValue($index * 10);
            $sheet->Cells['J' . $index]->PutValue($index * 10);
            $sheet->Cells['K' . $index]->PutValue($index * 10);
            $sheet->Cells['L' . $index]->PutValue($index * 10);
            $sheet->Cells['M' . $index]->PutValue($index * 10);
            $sheet->Cells['N' . $index]->PutValue($index * 10);
        }

        $contains = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $contains->AddArea(CellArea::CreateCellArea('A1', 'A10'));
        $containsRule = $contains[$contains->AddCondition(FormatConditionType::ContainsText)];
        $containsRule->Formula1 = 'error';
        $containsStyle = $containsRule->Style;
        $containsStyle->Pattern = FillPattern::Solid;
        $containsStyle->ForegroundColor = Color::FromArgb(255, 255, 235, 156);
        $containsStyle->Font->Color = Color::FromArgb(255, 156, 0, 6);
        $containsRule->Style = $containsStyle;

        $notContains = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $notContains->AddArea(CellArea::CreateCellArea('B1', 'B10'));
        $notContainsRule = $notContains[$notContains->AddCondition(FormatConditionType::NotContainsText)];
        $notContainsRule->Formula1 = 'warning';

        $beginsWith = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $beginsWith->AddArea(CellArea::CreateCellArea('C1', 'C10'));
        $beginsWithRule = $beginsWith[$beginsWith->AddCondition(FormatConditionType::BeginsWith)];
        $beginsWithRule->Formula1 = 'prefix';

        $endsWith = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $endsWith->AddArea(CellArea::CreateCellArea('D1', 'D10'));
        $endsWithRule = $endsWith[$endsWith->AddCondition(FormatConditionType::EndsWith)];
        $endsWithRule->Formula1 = 'suffix';

        $timePeriod = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $timePeriod->AddArea(CellArea::CreateCellArea('E1', 'E10'));
        $timePeriodRule = $timePeriod[$timePeriod->AddCondition(FormatConditionType::TimePeriod)];
        $timePeriodRule->TimePeriod = 'today';
        $timeStyle = $timePeriodRule->Style;
        $timeStyle->Pattern = FillPattern::Solid;
        $timeStyle->ForegroundColor = Color::FromArgb(255, 221, 235, 247);
        $timePeriodRule->Style = $timeStyle;

        $duplicate = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $duplicate->AddArea(CellArea::CreateCellArea('F1', 'F10'));
        $duplicate[$duplicate->AddCondition(FormatConditionType::DuplicateValues)]->Duplicate = true;

        $unique = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $unique->AddArea(CellArea::CreateCellArea('G1', 'G10'));
        $unique[$unique->AddCondition(FormatConditionType::UniqueValues)]->Duplicate = false;

        $top10 = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $top10->AddArea(CellArea::CreateCellArea('H1', 'H10'));
        $top10Rule = $top10[$top10->AddCondition(FormatConditionType::Top10)];
        $top10Rule->Top = true;
        $top10Rule->Percent = true;
        $top10Rule->Rank = 10;

        $bottom10 = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $bottom10->AddArea(CellArea::CreateCellArea('I1', 'I10'));
        $bottom10Rule = $bottom10[$bottom10->AddCondition(FormatConditionType::Bottom10)];
        $bottom10Rule->Top = false;
        $bottom10Rule->Rank = 10;

        $aboveAverage = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $aboveAverage->AddArea(CellArea::CreateCellArea('J1', 'J10'));
        $aboveAverageRule = $aboveAverage[$aboveAverage->AddCondition(FormatConditionType::AboveAverage)];
        $aboveAverageRule->Above = true;

        $belowAverage = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $belowAverage->AddArea(CellArea::CreateCellArea('K1', 'K10'));
        $belowAverageRule = $belowAverage[$belowAverage->AddCondition(FormatConditionType::BelowAverage)];
        $belowAverageRule->Above = false;
        $belowAverageRule->StandardDeviation = 1;

        $colorScale = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $colorScale->AddArea(CellArea::CreateCellArea('L1', 'L10'));
        $colorScaleRule = $colorScale[$colorScale->AddCondition(FormatConditionType::ColorScale)];
        $colorScaleRule->ColorScaleCount = 3;
        $colorScaleRule->MinColor = Color::FromArgb(255, 248, 105, 107);
        $colorScaleRule->MidColor = Color::FromArgb(255, 255, 235, 132);
        $colorScaleRule->MaxColor = Color::FromArgb(255, 99, 190, 123);

        $dataBar = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $dataBar->AddArea(CellArea::CreateCellArea('M1', 'M10'));
        $dataBarRule = $dataBar[$dataBar->AddCondition(FormatConditionType::DataBar)];
        $dataBarRule->BarColor = Color::FromArgb(255, 99, 142, 198);
        $dataBarRule->NegativeBarColor = Color::FromArgb(255, 255, 0, 0);
        $dataBarRule->ShowBorder = true;
        $dataBarRule->Direction = 'left-to-right';

        $iconSet = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $iconSet->AddArea(CellArea::CreateCellArea('N1', 'N10'));
        $iconSetRule = $iconSet[$iconSet->AddCondition(FormatConditionType::IconSet)];
        $iconSetRule->IconSetType = '4Arrows';
        $iconSetRule->ReverseIcons = true;
        $iconSetRule->ShowIconOnly = true;

        return $workbook;
    }

    public static function AssertAdvancedConditionalFormattings(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets['Advanced Conditional Formatting'];
        AssertEx::Equal(14, $sheet->ConditionalFormattings->Count);

        AssertEx::Equal('error', self::GetCollectionByArea($sheet->ConditionalFormattings, 'A1:A10')[0]->Formula1);
        AssertEx::Equal(FormatConditionType::ContainsText, self::GetCollectionByArea($sheet->ConditionalFormattings, 'A1:A10')[0]->Type);
        AssertEx::Equal(FillPattern::Solid, self::GetCollectionByArea($sheet->ConditionalFormattings, 'A1:A10')[0]->Style->Pattern);

        AssertEx::Equal(FormatConditionType::NotContainsText, self::GetCollectionByArea($sheet->ConditionalFormattings, 'B1:B10')[0]->Type);
        AssertEx::Equal('warning', self::GetCollectionByArea($sheet->ConditionalFormattings, 'B1:B10')[0]->Formula1);

        AssertEx::Equal(FormatConditionType::BeginsWith, self::GetCollectionByArea($sheet->ConditionalFormattings, 'C1:C10')[0]->Type);
        AssertEx::Equal('prefix', self::GetCollectionByArea($sheet->ConditionalFormattings, 'C1:C10')[0]->Formula1);

        AssertEx::Equal(FormatConditionType::EndsWith, self::GetCollectionByArea($sheet->ConditionalFormattings, 'D1:D10')[0]->Type);
        AssertEx::Equal('suffix', self::GetCollectionByArea($sheet->ConditionalFormattings, 'D1:D10')[0]->Formula1);

        AssertEx::Equal(FormatConditionType::TimePeriod, self::GetCollectionByArea($sheet->ConditionalFormattings, 'E1:E10')[0]->Type);
        AssertEx::Equal('today', self::GetCollectionByArea($sheet->ConditionalFormattings, 'E1:E10')[0]->TimePeriod);

        AssertEx::Equal(FormatConditionType::DuplicateValues, self::GetCollectionByArea($sheet->ConditionalFormattings, 'F1:F10')[0]->Type);
        AssertEx::True(self::GetCollectionByArea($sheet->ConditionalFormattings, 'F1:F10')[0]->Duplicate);

        AssertEx::Equal(FormatConditionType::UniqueValues, self::GetCollectionByArea($sheet->ConditionalFormattings, 'G1:G10')[0]->Type);
        AssertEx::False(self::GetCollectionByArea($sheet->ConditionalFormattings, 'G1:G10')[0]->Duplicate);

        $topRule = self::GetCollectionByArea($sheet->ConditionalFormattings, 'H1:H10')[0];
        AssertEx::Equal(FormatConditionType::Top10, $topRule->Type);
        AssertEx::True($topRule->Top);
        AssertEx::True($topRule->Percent);
        AssertEx::Equal(10, $topRule->Rank);

        $bottomRule = self::GetCollectionByArea($sheet->ConditionalFormattings, 'I1:I10')[0];
        AssertEx::Equal(FormatConditionType::Bottom10, $bottomRule->Type);
        AssertEx::False($bottomRule->Top);
        AssertEx::Equal(10, $bottomRule->Rank);

        AssertEx::Equal(FormatConditionType::AboveAverage, self::GetCollectionByArea($sheet->ConditionalFormattings, 'J1:J10')[0]->Type);
        AssertEx::True(self::GetCollectionByArea($sheet->ConditionalFormattings, 'J1:J10')[0]->Above);

        $belowRule = self::GetCollectionByArea($sheet->ConditionalFormattings, 'K1:K10')[0];
        AssertEx::Equal(FormatConditionType::BelowAverage, $belowRule->Type);
        AssertEx::False($belowRule->Above);
        AssertEx::Equal(1, $belowRule->StandardDeviation);

        $colorScaleRule = self::GetCollectionByArea($sheet->ConditionalFormattings, 'L1:L10')[0];
        AssertEx::Equal(FormatConditionType::ColorScale, $colorScaleRule->Type);
        AssertEx::Equal(3, $colorScaleRule->ColorScaleCount);
        AssertEx::Equal(Color::FromArgb(255, 248, 105, 107), $colorScaleRule->MinColor);
        AssertEx::Equal(Color::FromArgb(255, 255, 235, 132), $colorScaleRule->MidColor);
        AssertEx::Equal(Color::FromArgb(255, 99, 190, 123), $colorScaleRule->MaxColor);

        $dataBarRule = self::GetCollectionByArea($sheet->ConditionalFormattings, 'M1:M10')[0];
        AssertEx::Equal(FormatConditionType::DataBar, $dataBarRule->Type);
        AssertEx::Equal(Color::FromArgb(255, 99, 142, 198), $dataBarRule->BarColor);

        $iconSetRule = self::GetCollectionByArea($sheet->ConditionalFormattings, 'N1:N10')[0];
        AssertEx::Equal(FormatConditionType::IconSet, $iconSetRule->Type);
        AssertEx::Equal('4Arrows', $iconSetRule->IconSetType);
        AssertEx::True($iconSetRule->ReverseIcons);
        AssertEx::True($iconSetRule->ShowIconOnly);
    }
}

final class DefinedNameScenarioFactory
{
    public static function CreateDefinedNamesWorkbook()
    {
        $workbook = PageSetupScenarioFactory::CreatePageSetupWorkbook();
        $workbook->Worksheets->Add('Scoped Sheet');
        $workbook->Worksheets[1]->Cells['B2']->PutValue(5);

        $global = $workbook->DefinedNames[$workbook->DefinedNames->Add('GlobalRange', "='Print Sheet'!\$A$1:\$C$10")];
        $global->Hidden = true;
        $global->Comment = 'Primary range';

        $local = $workbook->DefinedNames[$workbook->DefinedNames->Add('LocalCell', "'Scoped Sheet'!\$B$2", 1)];
        $local->Comment = 'Scoped name';

        return $workbook;
    }

    public static function AssertDefinedNames(Workbook $workbook)
    {
        PageSetupScenarioFactory::AssertPageSetup($workbook);
        AssertEx::Equal(2, $workbook->DefinedNames->Count);
        AssertEx::Equal('5', $workbook->Worksheets[1]->Cells['B2']->StringValue);

        $global = $workbook->DefinedNames[0];
        AssertEx::Equal('GlobalRange', $global->Name);
        AssertEx::Equal("'Print Sheet'!\$A\$1:\$C\$10", $global->Formula);
        AssertEx::Null($global->LocalSheetIndex);
        AssertEx::True($global->Hidden);
        AssertEx::Equal('Primary range', $global->Comment);

        $local = $workbook->DefinedNames[1];
        AssertEx::Equal('LocalCell', $local->Name);
        AssertEx::Equal("'Scoped Sheet'!\$B\$2", $local->Formula);
        AssertEx::Equal(1, $local->LocalSheetIndex ?? -1);
        AssertEx::False($local->Hidden);
        AssertEx::Equal('Scoped name', $local->Comment);
    }
}
