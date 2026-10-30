<?php

namespace Aspose\Cells_FOSS\Samples\Basic;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\Cell;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'cell-data-roundtrip.xlsx';
        $timestamp = new \DateTime('2024-05-06 07:08:09', new \DateTimeZone('UTC'));

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells['A1']->PutValue('Hello');
        $sheet->Cells['B1']->PutValue(123);
        $sheet->Cells['C1']->PutValue(true);
        $sheet->Cells['D1']->PutValue(12.5);
        $sheet->Cells['E1']->PutValue($timestamp);
        $sheet->Cells['F1']->PutValue(10);
        $sheet->Cells['G1']->PutValue(20);
        $sheet->Cells['G1']->Formula = '=F1*2';
        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets[0];

        echo $loadedSheet->Cells['A1']->StringValue . PHP_EOL;
        echo self::GetValueTypeName($loadedSheet->Cells['B1']) . ':' . $loadedSheet->Cells['B1']->StringValue . PHP_EOL;
        echo self::GetValueTypeName($loadedSheet->Cells['C1']) . ':' . $loadedSheet->Cells['C1']->StringValue . PHP_EOL;
        echo self::GetValueTypeName($loadedSheet->Cells['D1']) . ':' . $loadedSheet->Cells['D1']->StringValue . PHP_EOL;
        echo self::GetValueTypeName($loadedSheet->Cells['E1']) . ':' . $loadedSheet->Cells['E1']->StringValue . PHP_EOL;
        echo $loadedSheet->Cells['G1']->Formula . ' -> ' . $loadedSheet->Cells['G1']->StringValue . PHP_EOL;
    }

    private static function GetValueTypeName(Cell $cell)
    {
        if ($cell->Value === null) {
            return '';
        }

        return get_debug_type($cell->Value);
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
