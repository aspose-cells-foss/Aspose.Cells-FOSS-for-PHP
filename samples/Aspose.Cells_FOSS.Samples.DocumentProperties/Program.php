<?php

namespace Aspose\Cells_FOSS\Samples\DocumentProperties;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'document-properties-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Sample';

        $sheet->Cells['A1']->PutValue('Title');
        $sheet->Cells['B1']->PutValue('Document Properties Sample');

        $properties = $workbook->DocumentProperties;
        $properties->Title = 'Annual Sales Report 2024';
        $properties->Subject = 'Financial Performance Analysis';
        $properties->Author = 'Finance Department';
        $properties->Keywords = 'sales, finance, 2024, report';
        $properties->Comments = 'This document contains quarterly sales data and projections';
        $properties->Category = 'Financial Reports';
        $properties->Company = 'Acme Corporation';
        $properties->Manager = 'Jane Smith';

        $properties->Core->Creator = 'Finance Department';
        $properties->Core->LastModifiedBy = 'Finance Department';
        $properties->Core->Created = new \DateTime('now', new \DateTimeZone('UTC'));

        $properties->Extended->Company = 'Acme Corporation';
        $properties->Extended->Manager = 'Jane Smith';

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedProperties = $loaded->DocumentProperties;

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Title: ' . $loadedProperties->Title . PHP_EOL;
        echo 'Subject: ' . $loadedProperties->Subject . PHP_EOL;
        echo 'Author: ' . $loadedProperties->Author . PHP_EOL;
        echo 'Keywords: ' . $loadedProperties->Keywords . PHP_EOL;
        echo 'Comments: ' . $loadedProperties->Comments . PHP_EOL;
        echo 'Category: ' . $loadedProperties->Category . PHP_EOL;
        echo 'Company: ' . $loadedProperties->Company . PHP_EOL;
        echo 'Manager: ' . $loadedProperties->Manager . PHP_EOL;
        echo 'Core Creator: ' . $loadedProperties->Core->Creator . PHP_EOL;
        echo 'Core LastModifiedBy: ' . $loadedProperties->Core->LastModifiedBy . PHP_EOL;
        echo 'Extended Company: ' . $loadedProperties->Extended->Company . PHP_EOL;
        echo 'Extended Manager: ' . $loadedProperties->Extended->Manager . PHP_EOL;
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
