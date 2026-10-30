<?php

namespace Aspose\Cells_FOSS\Samples\Loading;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\LoadOptions;
use Aspose\Cells_FOSS\Workbook;
use Aspose\Cells_FOSS\WorkbookLoadException;

final class Program
{
    public static function Main()
    {
        $options = new LoadOptions();
        $options->TryRepairPackage = true;
        $options->TryRepairXml = true;

        try {
            $inputPath = self::ResolveInputPath();
            new Workbook($inputPath, $options);
            echo 'Loaded: ' . $inputPath . PHP_EOL;
        } catch (WorkbookLoadException $exception) {
            echo $exception->getMessage() . PHP_EOL;
        }
    }

    private static function ResolveInputPath()
    {
        $directory = getcwd();
        while ($directory !== false && $directory !== '') {
            $inputPath = $directory . DIRECTORY_SEPARATOR . 'samples' . DIRECTORY_SEPARATOR . 'files' . DIRECTORY_SEPARATOR . 'sample.xlsx';
            if (is_file($inputPath)) {
                return $inputPath;
            }
            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }
            $directory = $parent;
        }

        throw new \RuntimeException('Could not find samples/files/sample.xlsx.');
    }
}

Program::Main();
