<?php

namespace Aspose\Cells_FOSS\Testing;

use Aspose\Cells_FOSS\TableStyleType;
use Aspose\Cells_FOSS\TotalsCalculation;
use Aspose\Cells_FOSS\Workbook;

final class ListObjectScenarioFactory
{
    public static function CreateTableWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'SalesData';

        $sheet->Cells['A1']->PutValue('Product');
        $sheet->Cells['B1']->PutValue('Region');
        $sheet->Cells['C1']->PutValue('Revenue');
        $sheet->Cells['A2']->PutValue('Widget');
        $sheet->Cells['B2']->PutValue('North');
        $sheet->Cells['C2']->PutValue(1200);
        $sheet->Cells['A3']->PutValue('Gadget');
        $sheet->Cells['B3']->PutValue('South');
        $sheet->Cells['C3']->PutValue(850);
        $sheet->Cells['A4']->PutValue('Doohickey');
        $sheet->Cells['B4']->PutValue('East');
        $sheet->Cells['C4']->PutValue(2100);

        $index = $sheet->ListObjects->Add('A1', 'C4', true);
        $table = $sheet->ListObjects[$index];
        $table->TableStyleType = TableStyleType::TableStyleMedium2;
        $table->ShowTableStyleRowStripes = true;
        $table->ShowTableStyleFirstColumn = false;
        $table->ShowTableStyleLastColumn = false;
        $table->ShowTableStyleColumnStripes = false;
        $table->ShowTotals = true;
        $table->ListColumns[2]->TotalsCalculation = TotalsCalculation::Sum;
        $table->ListColumns[0]->TotalsRowLabel = 'Total';

        return $workbook;
    }

    public static function AssertTableWorkbook(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets[0];
        AssertEx::Equal(1, $sheet->ListObjects->Count);

        $table = $sheet->ListObjects[0];
        AssertEx::Equal('Table1', $table->DisplayName);
        AssertEx::Equal(0, $table->StartRow);
        AssertEx::Equal(0, $table->StartColumn);
        AssertEx::Equal(4, $table->EndRow);
        AssertEx::Equal(2, $table->EndColumn);
        AssertEx::True($table->ShowHeaderRow);
        AssertEx::True($table->ShowTotals);
        AssertEx::Equal(TableStyleType::TableStyleMedium2, $table->TableStyleType);
        AssertEx::Equal('TableStyleMedium2', $table->TableStyleName);
        AssertEx::True($table->ShowTableStyleRowStripes);
        AssertEx::False($table->ShowTableStyleFirstColumn);
        AssertEx::False($table->ShowTableStyleLastColumn);
        AssertEx::False($table->ShowTableStyleColumnStripes);
        AssertEx::Equal(3, $table->ListColumns->Count);
        AssertEx::Equal('Product', $table->ListColumns[0]->Name);
        AssertEx::Equal('Region', $table->ListColumns[1]->Name);
        AssertEx::Equal('Revenue', $table->ListColumns[2]->Name);
        AssertEx::Equal(TotalsCalculation::Sum, $table->ListColumns[2]->TotalsCalculation);
        AssertEx::Equal(TotalsCalculation::None, $table->ListColumns[0]->TotalsCalculation);
        AssertEx::Equal('Total', $table->ListColumns[0]->TotalsRowLabel);
    }
}
