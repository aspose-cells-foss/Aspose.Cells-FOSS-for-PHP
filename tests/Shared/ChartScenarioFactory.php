<?php

namespace Aspose\Cells_FOSS\Testing;

use Aspose\Cells_FOSS\Workbook;

final class ChartScenarioFactory
{
    private static $QuarterLabels = array('Q1', 'Q2', 'Q3', 'Q4', 'Q5');
    private static $QuarterValues = array(10, 30, 20, 50, 40);

    public static function CreateChartWorkbook($type, $sheetName = 'Sheet1')
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = $sheetName;

        $sheet->Cells['A1']->PutValue(self::$QuarterLabels[0]);
        $sheet->Cells['B1']->PutValue(self::$QuarterValues[0]);
        $sheet->Cells['A2']->PutValue(self::$QuarterLabels[1]);
        $sheet->Cells['B2']->PutValue(self::$QuarterValues[1]);
        $sheet->Cells['A3']->PutValue(self::$QuarterLabels[2]);
        $sheet->Cells['B3']->PutValue(self::$QuarterValues[2]);
        $sheet->Cells['A4']->PutValue(self::$QuarterLabels[3]);
        $sheet->Cells['B4']->PutValue(self::$QuarterValues[3]);
        $sheet->Cells['A5']->PutValue(self::$QuarterLabels[4]);
        $sheet->Cells['B5']->PutValue(self::$QuarterValues[4]);

        $dataRange = $sheetName . '!$B$1:$B$5';
        $sheet->Charts->Add($type, $dataRange, 7, 0, 22, 10);
        return $workbook;
    }

    public static function AssertChartWorkbook(Workbook $workbook, $expectedType)
    {
        $sheet = $workbook->Worksheets[0];
        AssertEx::Equal(1, $sheet->Charts->Count, 'Expected exactly 1 chart');
        $chart = $sheet->Charts[0];
        AssertEx::Equal($expectedType, $chart->ChartType, 'Wrong chart type after round-trip');
        AssertEx::True(!($chart->Name === null || $chart->Name === ''), 'Chart name should not be empty');
        AssertEx::Equal(7, $chart->UpperLeftRow, 'Wrong UpperLeftRow');
        AssertEx::Equal(0, $chart->UpperLeftColumn, 'Wrong UpperLeftColumn');
        AssertEx::Equal(22, $chart->LowerRightRow, 'Wrong LowerRightRow');
        AssertEx::Equal(10, $chart->LowerRightColumn, 'Wrong LowerRightColumn');
    }
}
