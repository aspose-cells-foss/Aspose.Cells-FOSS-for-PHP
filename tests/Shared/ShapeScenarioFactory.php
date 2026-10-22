<?php

namespace Aspose\Cells_FOSS\Testing;

use Aspose\Cells_FOSS\AutoShapeType;
use Aspose\Cells_FOSS\Workbook;

final class ShapeScenarioFactory
{
    public static function CreateShapeWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Shapes';

        $sheet->Cells['A1']->PutValue('Drawing Objects Demo');

        $rectIndex = $sheet->Shapes->Add(2, 0, 6, 4, AutoShapeType::Rectangle);
        $sheet->Shapes[$rectIndex]->Name = 'MyRectangle';

        $arrowIndex = $sheet->Shapes->Add(8, 0, 11, 4, AutoShapeType::RightArrow);
        $sheet->Shapes[$arrowIndex]->Name = 'MyRightArrow';

        $ellipseIndex = $sheet->Shapes->Add(13, 0, 17, 4, AutoShapeType::Ellipse);
        $sheet->Shapes[$ellipseIndex]->Name = 'MyEllipse';

        $starIndex = $sheet->Shapes->Add(2, 5, 7, 9, AutoShapeType::Star5Point);
        $sheet->Shapes[$starIndex]->Name = 'MyStar5';

        $downArrowIndex = $sheet->Shapes->Add(9, 5, 12, 9, AutoShapeType::DownArrow);
        $sheet->Shapes[$downArrowIndex]->Name = 'MyDownArrow';

        $star12Index = $sheet->Shapes->Add(14, 5, 19, 9, AutoShapeType::Star12Point);
        $sheet->Shapes[$star12Index]->Name = 'MyStar12';

        return $workbook;
    }

    public static function AssertShapeWorkbook(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets[0];
        AssertEx::Equal(6, $sheet->Shapes->Count);

        $rect = $sheet->Shapes[0];
        AssertEx::Equal('MyRectangle', $rect->Name);
        AssertEx::Equal(2, $rect->UpperLeftRow);
        AssertEx::Equal(0, $rect->UpperLeftColumn);
        AssertEx::Equal(6, $rect->LowerRightRow);
        AssertEx::Equal(4, $rect->LowerRightColumn);
        AssertEx::Equal(AutoShapeType::Rectangle, $rect->AutoShapeType);

        $arrow = $sheet->Shapes[1];
        AssertEx::Equal('MyRightArrow', $arrow->Name);
        AssertEx::Equal(AutoShapeType::RightArrow, $arrow->AutoShapeType);

        $ellipse = $sheet->Shapes[2];
        AssertEx::Equal('MyEllipse', $ellipse->Name);
        AssertEx::Equal(AutoShapeType::Ellipse, $ellipse->AutoShapeType);

        $star = $sheet->Shapes[3];
        AssertEx::Equal('MyStar5', $star->Name);
        AssertEx::Equal(AutoShapeType::Star5Point, $star->AutoShapeType);

        $downArrow = $sheet->Shapes[4];
        AssertEx::Equal('MyDownArrow', $downArrow->Name);
        AssertEx::Equal(AutoShapeType::DownArrow, $downArrow->AutoShapeType);

        $star12 = $sheet->Shapes[5];
        AssertEx::Equal('MyStar12', $star12->Name);
        AssertEx::Equal(AutoShapeType::Star12Point, $star12->AutoShapeType);
    }
}
