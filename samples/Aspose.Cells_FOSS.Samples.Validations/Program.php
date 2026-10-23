<?php

namespace Aspose\Cells_FOSS\Samples\Validations;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\Validation;
use Aspose\Cells_FOSS\ValidationAlertType;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'validations-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Validation Sheet';

        $sheet->Cells['A1']->PutValue('Open');
        $sheet->Cells['B2']->PutValue(5);
        $sheet->Cells['C3']->PutValue(7);
        $sheet->Cells['E2']->PutValue(8);
        $sheet->Cells['G1']->PutValue('ABCDE');

        $listValidationIndex = $sheet->Validations->Add(CellArea::CreateCellArea('A1', 'A3'));
        $listValidation = $sheet->Validations[$listValidationIndex];
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

        $decimalValidationIndex = $sheet->Validations->Add(CellArea::CreateCellArea('B2', 'C3'));
        $decimalValidation = $sheet->Validations[$decimalValidationIndex];
        $decimalValidation->Type = ValidationType::Decimal;
        $decimalValidation->Operator = OperatorType::Between;
        $decimalValidation->Formula1 = '1.5';
        $decimalValidation->Formula2 = '9.5';
        $decimalValidation->ErrorTitle = 'Range';
        $decimalValidation->ErrorMessage = 'Enter 1.5-9.5';
        $decimalValidation->ShowError = true;
        $decimalValidation->AddArea(CellArea::CreateCellArea('E2', 'E3'));

        $customValidationIndex = $sheet->Validations->Add(CellArea::CreateCellArea('G1', 'G1'));
        $customValidation = $sheet->Validations[$customValidationIndex];
        $customValidation->Type = ValidationType::Custom;
        $customValidation->AlertStyle = ValidationAlertType::Warning;
        $customValidation->Formula1 = 'LEN(G1)<=5';
        $customValidation->ShowInput = true;
        $customValidation->InputTitle = 'Code';
        $customValidation->InputMessage = 'Up to 5 chars';

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Validation Sheet'];
        $a1Validation = self::RequireValidation($loadedSheet->Validations->GetValidationInCell(0, 0), 'A1');
        $c2Validation = self::RequireValidation($loadedSheet->Validations->GetValidationInCell(1, 2), 'C2');
        $g1Validation = self::RequireValidation($loadedSheet->Validations->GetValidationInCell(0, 6), 'G1');

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Validation count: ' . $loadedSheet->Validations->Count . PHP_EOL;
        echo 'A1 validation: ' . self::ValidationTypeName($a1Validation->Type) . ' / ' . $a1Validation->Formula1 . PHP_EOL;
        echo 'C2 validation: ' . self::ValidationTypeName($c2Validation->Type) . ' / ' . $c2Validation->Formula1 . ' to ' . $c2Validation->Formula2 . PHP_EOL;
        echo 'G1 validation: ' . self::ValidationTypeName($g1Validation->Type) . PHP_EOL;
    }

    private static function RequireValidation($validation, $cellName)
    {
        if ($validation === null) {
            throw new \RuntimeException('Expected validation at ' . $cellName . '.');
        }

        return $validation;
    }

    private static function ValidationTypeName($value)
    {
        switch ($value) {
            case ValidationType::AnyValue:
                return 'AnyValue';
            case ValidationType::WholeNumber:
                return 'WholeNumber';
            case ValidationType::Decimal:
                return 'Decimal';
            case ValidationType::List:
                return 'List';
            case ValidationType::Date:
                return 'Date';
            case ValidationType::Time:
                return 'Time';
            case ValidationType::TextLength:
                return 'TextLength';
            case ValidationType::Custom:
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
