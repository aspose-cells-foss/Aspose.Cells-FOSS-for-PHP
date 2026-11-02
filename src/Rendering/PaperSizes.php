<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PaperSizes
{
    private function __construct()
    {
    }

    public static function GetDimensionsPoints($paperSizeCode, &$widthPt, &$heightPt)
    {
        switch ($paperSizeCode) {
            case 1:
                $widthIn = 8.5;
                $heightIn = 11.0;
                break;
            case 3:
                $widthIn = 11.0;
                $heightIn = 17.0;
                break;
            case 5:
                $widthIn = 8.5;
                $heightIn = 14.0;
                break;
            case 8:
                $widthIn = 11.69;
                $heightIn = 16.54;
                break;
            case 9:
                $widthIn = 210.0 / 25.4;
                $heightIn = 297.0 / 25.4;
                break;
            case 11:
                $widthIn = 5.83;
                $heightIn = 8.27;
                break;
            case 13:
                $widthIn = 7.17;
                $heightIn = 10.12;
                break;
            default:
                $widthIn = 210.0 / 25.4;
                $heightIn = 297.0 / 25.4;
                break;
        }
        $widthPt = RenderUnits::InchesToPoints($widthIn);
        $heightPt = RenderUnits::InchesToPoints($heightIn);
    }
}
