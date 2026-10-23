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
            new Workbook('sample.xlsx', $options);
        } catch (WorkbookLoadException $exception) {
            echo $exception->getMessage() . PHP_EOL;
        }
    }
}

Program::Main();
