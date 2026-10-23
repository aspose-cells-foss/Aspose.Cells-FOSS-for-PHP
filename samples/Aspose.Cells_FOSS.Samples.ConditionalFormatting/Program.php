<?php

namespace Aspose\Cells_FOSS\Samples\ConditionalFormatting;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'conditional-formatting-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Conditional Formatting';

        for ($index = 0; $index < 10; $index++) {
            $sheet->Cells[[$index, 0]]->PutValue($index + 1);
            $sheet->Cells[[$index, 1]]->PutValue(($index + 1) * 10);
            $sheet->Cells[[$index, 2]]->PutValue(($index + 1) * 10);
            $sheet->Cells[[$index, 3]]->PutValue(($index + 1) * 10);
            $sheet->Cells[[$index, 4]]->PutValue(($index + 1) * 10);
        }

        $betweenCollection = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $betweenCollection->AddArea(CellArea::CreateCellArea('A1', 'A10'));
        $betweenRule = $betweenCollection[$betweenCollection->AddCondition(FormatConditionType::CellValue, OperatorType::Between, '3', '7')];
        $betweenRule->Priority = 1;
        $betweenRule->StopIfTrue = true;
        $betweenStyle = $betweenRule->Style;
        $betweenStyle->Pattern = FillPattern::Solid;
        $betweenStyle->ForegroundColor = Color::FromArgb(255, 255, 199, 206);
        $betweenStyle->Font->IsBold = true;
        $betweenStyle->Font->Color = Color::FromArgb(255, 156, 0, 6);
        $betweenRule->Style = $betweenStyle;

        $expressionCollection = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $expressionCollection->AddArea(CellArea::CreateCellArea('B1', 'B10'));
        $expressionRule = $expressionCollection[$expressionCollection->AddCondition(FormatConditionType::Expression, OperatorType::None, 'MOD(B1,20)=0', '')];
        $expressionRule->Priority = 2;
        $expressionStyle = $expressionRule->Style;
        $expressionStyle->Font->IsItalic = true;
        $expressionStyle->Font->Color = Color::FromArgb(255, 0, 0, 255);
        $expressionRule->Style = $expressionStyle;

        $colorScaleCollection = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $colorScaleCollection->AddArea(CellArea::CreateCellArea('C1', 'C10'));
        $colorScaleRule = $colorScaleCollection[$colorScaleCollection->AddCondition(FormatConditionType::ColorScale)];
        $colorScaleRule->ColorScaleCount = 3;
        $colorScaleRule->MinColor = Color::FromArgb(255, 248, 105, 107);
        $colorScaleRule->MidColor = Color::FromArgb(255, 255, 235, 132);
        $colorScaleRule->MaxColor = Color::FromArgb(255, 99, 190, 123);

        $dataBarCollection = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $dataBarCollection->AddArea(CellArea::CreateCellArea('D1', 'D10'));
        $dataBarRule = $dataBarCollection[$dataBarCollection->AddCondition(FormatConditionType::DataBar)];
        $dataBarRule->BarColor = Color::FromArgb(255, 99, 142, 198);
        $dataBarRule->NegativeBarColor = Color::FromArgb(255, 255, 0, 0);
        $dataBarRule->ShowBorder = true;
        $dataBarRule->Direction = 'left-to-right';

        $iconSetCollection = $sheet->ConditionalFormattings[$sheet->ConditionalFormattings->Add()];
        $iconSetCollection->AddArea(CellArea::CreateCellArea('E1', 'E10'));
        $iconSetRule = $iconSetCollection[$iconSetCollection->AddCondition(FormatConditionType::IconSet)];
        $iconSetRule->IconSetType = '4Arrows';
        $iconSetRule->ReverseIcons = true;
        $iconSetRule->ShowIconOnly = true;

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Conditional Formatting'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Conditional formatting collections: ' . $loadedSheet->ConditionalFormattings->Count . PHP_EOL;
        echo 'A-column rule type: ' . self::FormatConditionTypeName($loadedSheet->ConditionalFormattings[0][0]->Type) . PHP_EOL;
        echo 'C-column rule type: ' . self::FormatConditionTypeName($loadedSheet->ConditionalFormattings[2][0]->Type) . PHP_EOL;
        echo 'E-column icon set: ' . $loadedSheet->ConditionalFormattings[4][0]->IconSetType . PHP_EOL;
    }

    private static function FormatConditionTypeName($value)
    {
        switch ($value) {
            case FormatConditionType::CellValue:
                return 'CellValue';
            case FormatConditionType::Expression:
                return 'Expression';
            case FormatConditionType::ContainsText:
                return 'ContainsText';
            case FormatConditionType::NotContainsText:
                return 'NotContainsText';
            case FormatConditionType::BeginsWith:
                return 'BeginsWith';
            case FormatConditionType::EndsWith:
                return 'EndsWith';
            case FormatConditionType::TimePeriod:
                return 'TimePeriod';
            case FormatConditionType::DuplicateValues:
                return 'DuplicateValues';
            case FormatConditionType::UniqueValues:
                return 'UniqueValues';
            case FormatConditionType::Top10:
                return 'Top10';
            case FormatConditionType::Bottom10:
                return 'Bottom10';
            case FormatConditionType::AboveAverage:
                return 'AboveAverage';
            case FormatConditionType::BelowAverage:
                return 'BelowAverage';
            case FormatConditionType::ColorScale:
                return 'ColorScale';
            case FormatConditionType::DataBar:
                return 'DataBar';
            case FormatConditionType::IconSet:
                return 'IconSet';
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
