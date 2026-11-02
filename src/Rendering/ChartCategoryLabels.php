<?php

namespace Aspose\Cells_FOSS\Rendering;

final class ChartCategoryLabels
{
    private function __construct()
    {
    }

    public static function Lines($label)
    {
        return explode("\n", str_replace("\r", "\n", str_replace("\r\n", "\n", $label !== null ? $label : '')));
    }

    public static function LineAdvance($font)
    {
        return max($font->LineHeightPt, $font->TextHeightPt);
    }

    public static function ExtraHeight($chart, $font)
    {
        $lines = 1;
        foreach ($chart->Categories as $label) {
            $lines = max($lines, count(self::Lines($label)));
        }

        return ($lines - 1) * self::LineAdvance($font);
    }
}
