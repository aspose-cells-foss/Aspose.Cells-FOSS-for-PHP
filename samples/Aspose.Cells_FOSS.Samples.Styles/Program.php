<?php

namespace Aspose\Cells_FOSS\Samples\Styles;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $workbook = new Workbook();
        $cell = $workbook->Worksheets[0]->Cells['A1'];
        $cell->PutValue('Styled');

        $style = $cell->GetStyle();
        $style->Font->IsBold = true;
        $style->Pattern = FillPattern::Solid;
        $style->ForegroundColor = Color::FromArgb(255, 241, 196, 15);
        $cell->SetStyle($style);

        echo $cell->StringValue . ' / Bold=' . self::FormatBool($cell->GetStyle()->Font->IsBold) . PHP_EOL;
    }

    private static function FormatBool($value)
    {
        return $value ? 'True' : 'False';
    }
}

Program::Main();
