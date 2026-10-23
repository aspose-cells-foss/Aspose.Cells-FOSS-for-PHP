<?php

namespace Aspose\Cells_FOSS\Samples\PageSetup;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\PageOrientationType;
use Aspose\Cells_FOSS\PaperSizeType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'page-setup-sample.xlsx';

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

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedPageSetup = $loaded->Worksheets[0]->PageSetup;

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Orientation: ' . self::PageOrientationTypeName($loadedPageSetup->Orientation) . PHP_EOL;
        echo 'Paper size: ' . self::PaperSizeTypeName($loadedPageSetup->PaperSize) . PHP_EOL;
        echo 'Print area: ' . $loadedPageSetup->PrintArea . PHP_EOL;
        echo 'Horizontal breaks: ' . implode(', ', $loadedPageSetup->HorizontalPageBreaks) . PHP_EOL;
        echo 'Vertical breaks: ' . implode(', ', $loadedPageSetup->VerticalPageBreaks) . PHP_EOL;
    }

    private static function PageOrientationTypeName($value)
    {
        switch ($value) {
            case PageOrientationType::Default:
                return 'Default';
            case PageOrientationType::Portrait:
                return 'Portrait';
            case PageOrientationType::Landscape:
                return 'Landscape';
            default:
                return (string)$value;
        }
    }

    private static function PaperSizeTypeName($value)
    {
        switch ($value) {
            case PaperSizeType::Default:
                return 'Default';
            case PaperSizeType::PaperLetter:
                return 'PaperLetter';
            case PaperSizeType::PaperTabloid:
                return 'PaperTabloid';
            case PaperSizeType::PaperLegal:
                return 'PaperLegal';
            case PaperSizeType::PaperStatement:
                return 'PaperStatement';
            case PaperSizeType::PaperExecutive:
                return 'PaperExecutive';
            case PaperSizeType::PaperA3:
                return 'PaperA3';
            case PaperSizeType::PaperA4:
                return 'PaperA4';
            case PaperSizeType::PaperA5:
                return 'PaperA5';
            case PaperSizeType::PaperB4:
                return 'PaperB4';
            case PaperSizeType::PaperB5:
                return 'PaperB5';
            default:
                return (string)$value;
        }
    }

    private static function ResolveOutputDirectory()
    {
        $directory = self::ResolveRepositoryRoot() . DIRECTORY_SEPARATOR . 'output' . DIRECTORY_SEPARATOR . 'samples';
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        return $directory;
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
}

Program::Main();
