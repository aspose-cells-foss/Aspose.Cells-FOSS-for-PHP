<?php

namespace Aspose\Cells_FOSS\Samples\WorksheetSettings;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\VisibilityType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'worksheet-settings-sample.xlsx';

        $workbook = new Workbook();
        $layout = $workbook->Worksheets[0];
        $layout->Name = 'Layout';
        $layout->VisibilityType = VisibilityType::Hidden;
        $layout->TabColor = Color::FromArgb(255, 34, 68, 102);
        $layout->ShowGridlines = false;
        $layout->ShowRowColumnHeaders = false;
        $layout->ShowZeros = false;
        $layout->RightToLeft = true;
        $layout->Zoom = 85;
        $layout->Protect();
        $layout->Protection->Objects = true;
        $layout->Protection->Scenarios = true;
        $layout->Protection->FormatCells = true;
        $layout->Protection->InsertRows = true;
        $layout->Protection->AutoFilter = true;
        $layout->Protection->SelectLockedCells = true;
        $layout->Protection->SelectUnlockedCells = true;

        $layout->Cells['A1']->PutValue('Merged');
        $layout->Cells['C4']->PutValue(99);
        $layout->Cells->Rows[1]->Height = 22.5;
        $layout->Cells->Rows[3]->IsHidden = true;
        $layout->Cells->Columns[0]->Width = 18.25;
        $layout->Cells->Columns[2]->IsHidden = true;
        $layout->Cells->Merge(0, 0, 2, 2);

        $visibleIndex = $workbook->Worksheets->Add('Visible');
        $visibleSheet = $workbook->Worksheets[$visibleIndex];
        $visibleSheet->Cells['A1']->PutValue('Visible');
        $workbook->Worksheets->ActiveSheetName = 'Visible';

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedLayout = $loaded->Worksheets['Layout'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Active sheet: ' . $loaded->Worksheets->ActiveSheetName . PHP_EOL;
        echo 'Layout visibility: ' . self::VisibilityTypeName($loadedLayout->VisibilityType) . PHP_EOL;
        echo 'Merged value: ' . $loadedLayout->Cells['A1']->StringValue . PHP_EOL;
        echo 'Row 2 height: ' . self::GetDoubleOrDefault($loadedLayout->Cells->Rows[1]->Height) . PHP_EOL;
        echo 'Column A width: ' . self::GetDoubleOrDefault($loadedLayout->Cells->Columns[0]->Width) . PHP_EOL;
        echo 'Merged regions: ' . count($loadedLayout->Cells->MergedCells) . PHP_EOL;
    }

    private static function GetDoubleOrDefault($value)
    {
        if ($value === null) {
            return 0.0;
        }

        return $value;
    }

    private static function VisibilityTypeName($value)
    {
        switch ($value) {
            case VisibilityType::Visible:
                return 'Visible';
            case VisibilityType::Hidden:
                return 'Hidden';
            case VisibilityType::VeryHidden:
                return 'VeryHidden';
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
