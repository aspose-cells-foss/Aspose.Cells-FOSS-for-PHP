<?php

namespace Aspose\Cells_FOSS\Samples\ListObjects;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\TableStyleType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'listobjects-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Data';

        $sheet->Cells['A1']->PutValue('Product');
        $sheet->Cells['B1']->PutValue('Category');
        $sheet->Cells['C1']->PutValue('Price');
        $sheet->Cells['D1']->PutValue('Quantity');

        $sheet->Cells['A2']->PutValue('Laptop');
        $sheet->Cells['B2']->PutValue('Electronics');
        $sheet->Cells['C2']->PutValue(999.99);
        $sheet->Cells['D2']->PutValue(50);

        $sheet->Cells['A3']->PutValue('Mouse');
        $sheet->Cells['B3']->PutValue('Electronics');
        $sheet->Cells['C3']->PutValue(29.99);
        $sheet->Cells['D3']->PutValue(200);

        $sheet->Cells['A4']->PutValue('Desk Chair');
        $sheet->Cells['B4']->PutValue('Furniture');
        $sheet->Cells['C4']->PutValue(149.99);
        $sheet->Cells['D4']->PutValue(75);

        $sheet->Cells['A5']->PutValue('Monitor');
        $sheet->Cells['B5']->PutValue('Electronics');
        $sheet->Cells['C5']->PutValue(349.99);
        $sheet->Cells['D5']->PutValue(100);

        $tableIndex = $sheet->ListObjects->Add('A1', 'D5', true);
        $table = $sheet->ListObjects[$tableIndex];
        $table->DisplayName = 'Products';
        $table->Comment = 'Product inventory table';
        $table->TableStyleType = TableStyleType::TableStyleMedium2;
        $table->ShowTableStyleFirstColumn = true;
        $table->ShowTableStyleLastColumn = true;
        $table->ShowTableStyleRowStripes = true;
        $table->ShowTotals = true;

        $sheet->Cells['A6']->PutValue('Total');
        $sheet->Cells['D6']->PutValue('=SUM(D2:D5)');

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Data'];
        $loadedTable = $loadedSheet->ListObjects['Products'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Table count: ' . $loadedSheet->ListObjects->Count . PHP_EOL;
        echo 'Table name: ' . $loadedTable->DisplayName . PHP_EOL;
        echo 'Table comment: ' . $loadedTable->Comment . PHP_EOL;
        echo 'Table style: ' . self::TableStyleTypeName($loadedTable->TableStyleType) . PHP_EOL;
        echo 'Show headers: ' . self::FormatBool($loadedTable->ShowHeaderRow) . PHP_EOL;
        echo 'Show totals: ' . self::FormatBool($loadedTable->ShowTotals) . PHP_EOL;
        echo 'Column count: ' . $loadedTable->ListColumns->Count . PHP_EOL;
        echo 'First column: ' . $loadedTable->ListColumns[0]->Name . PHP_EOL;
    }

    private static function FormatBool($value)
    {
        return $value ? 'True' : 'False';
    }

    private static function TableStyleTypeName($value)
    {
        switch ($value) {
            case TableStyleType::None:
                return 'None';
            case TableStyleType::TableStyleLight1:
                return 'TableStyleLight1';
            case TableStyleType::TableStyleLight2:
                return 'TableStyleLight2';
            case TableStyleType::TableStyleLight3:
                return 'TableStyleLight3';
            case TableStyleType::TableStyleLight4:
                return 'TableStyleLight4';
            case TableStyleType::TableStyleLight5:
                return 'TableStyleLight5';
            case TableStyleType::TableStyleLight6:
                return 'TableStyleLight6';
            case TableStyleType::TableStyleLight7:
                return 'TableStyleLight7';
            case TableStyleType::TableStyleLight8:
                return 'TableStyleLight8';
            case TableStyleType::TableStyleLight9:
                return 'TableStyleLight9';
            case TableStyleType::TableStyleLight10:
                return 'TableStyleLight10';
            case TableStyleType::TableStyleLight11:
                return 'TableStyleLight11';
            case TableStyleType::TableStyleLight12:
                return 'TableStyleLight12';
            case TableStyleType::TableStyleLight13:
                return 'TableStyleLight13';
            case TableStyleType::TableStyleLight14:
                return 'TableStyleLight14';
            case TableStyleType::TableStyleLight15:
                return 'TableStyleLight15';
            case TableStyleType::TableStyleLight16:
                return 'TableStyleLight16';
            case TableStyleType::TableStyleLight17:
                return 'TableStyleLight17';
            case TableStyleType::TableStyleLight18:
                return 'TableStyleLight18';
            case TableStyleType::TableStyleLight19:
                return 'TableStyleLight19';
            case TableStyleType::TableStyleLight20:
                return 'TableStyleLight20';
            case TableStyleType::TableStyleLight21:
                return 'TableStyleLight21';
            case TableStyleType::TableStyleMedium1:
                return 'TableStyleMedium1';
            case TableStyleType::TableStyleMedium2:
                return 'TableStyleMedium2';
            case TableStyleType::TableStyleMedium3:
                return 'TableStyleMedium3';
            case TableStyleType::TableStyleMedium4:
                return 'TableStyleMedium4';
            case TableStyleType::TableStyleMedium5:
                return 'TableStyleMedium5';
            case TableStyleType::TableStyleMedium6:
                return 'TableStyleMedium6';
            case TableStyleType::TableStyleMedium7:
                return 'TableStyleMedium7';
            case TableStyleType::TableStyleMedium8:
                return 'TableStyleMedium8';
            case TableStyleType::TableStyleMedium9:
                return 'TableStyleMedium9';
            case TableStyleType::TableStyleMedium10:
                return 'TableStyleMedium10';
            case TableStyleType::TableStyleMedium11:
                return 'TableStyleMedium11';
            case TableStyleType::TableStyleMedium12:
                return 'TableStyleMedium12';
            case TableStyleType::TableStyleMedium13:
                return 'TableStyleMedium13';
            case TableStyleType::TableStyleMedium14:
                return 'TableStyleMedium14';
            case TableStyleType::TableStyleMedium15:
                return 'TableStyleMedium15';
            case TableStyleType::TableStyleMedium16:
                return 'TableStyleMedium16';
            case TableStyleType::TableStyleMedium17:
                return 'TableStyleMedium17';
            case TableStyleType::TableStyleMedium18:
                return 'TableStyleMedium18';
            case TableStyleType::TableStyleMedium19:
                return 'TableStyleMedium19';
            case TableStyleType::TableStyleMedium20:
                return 'TableStyleMedium20';
            case TableStyleType::TableStyleMedium21:
                return 'TableStyleMedium21';
            case TableStyleType::TableStyleMedium22:
                return 'TableStyleMedium22';
            case TableStyleType::TableStyleMedium23:
                return 'TableStyleMedium23';
            case TableStyleType::TableStyleMedium24:
                return 'TableStyleMedium24';
            case TableStyleType::TableStyleMedium25:
                return 'TableStyleMedium25';
            case TableStyleType::TableStyleMedium26:
                return 'TableStyleMedium26';
            case TableStyleType::TableStyleMedium27:
                return 'TableStyleMedium27';
            case TableStyleType::TableStyleMedium28:
                return 'TableStyleMedium28';
            case TableStyleType::TableStyleDark1:
                return 'TableStyleDark1';
            case TableStyleType::TableStyleDark2:
                return 'TableStyleDark2';
            case TableStyleType::TableStyleDark3:
                return 'TableStyleDark3';
            case TableStyleType::TableStyleDark4:
                return 'TableStyleDark4';
            case TableStyleType::TableStyleDark5:
                return 'TableStyleDark5';
            case TableStyleType::TableStyleDark6:
                return 'TableStyleDark6';
            case TableStyleType::TableStyleDark7:
                return 'TableStyleDark7';
            case TableStyleType::TableStyleDark8:
                return 'TableStyleDark8';
            case TableStyleType::TableStyleDark9:
                return 'TableStyleDark9';
            case TableStyleType::TableStyleDark10:
                return 'TableStyleDark10';
            case TableStyleType::TableStyleDark11:
                return 'TableStyleDark11';
            case TableStyleType::Custom:
                return 'Custom';
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
