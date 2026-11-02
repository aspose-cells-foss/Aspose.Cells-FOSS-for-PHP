<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKRect;

final class LineChartAutomaticLayout
{
    private function __construct()
    {
    }

    public static function Applies($chart)
    {
        return self::IsAutomaticLine($chart) && $chart->HasDateCategoryAxis;
    }

    public static function AppliesToCategoryAxis($chart)
    {
        return self::IsAutomaticLine($chart) && !$chart->HasDateCategoryAxis;
    }

    private static function IsAutomaticLine($chart)
    {
        return $chart !== null && $chart->Kind === ChartKind::Line && !$chart->IsLine3D
            && $chart->CategoryAxisAutomaticTextRotation;
    }

    public static function Plot($area, $chart, $labelWidth, $rotation, $fallback)
    {
        $horizontalCategories = self::AppliesToCategoryAxis($chart) && abs($rotation) < 0.1;
        if ((!self::Applies($chart) && !$horizontalCategories) || $chart->HasManualPlot
            || (!$horizontalCategories && abs($rotation) < 0.1)
            || $chart->Title === null || $chart->LegendPosition !== 'b') {
            return $fallback;
        }
        // Text bands stop growing once the fixed-size title, labels and legend fit.
        // Taller charts should give their extra height to the inner plot.
        $topReserve = min($area->Height * 0.185, 36.0);
        $bottomReserve = $horizontalCategories
            ? min($area->Height * 0.239, 46.5)
            : min($area->Height * 0.318, 62.0);
        return new SKRect($area->Left + max($area->Width * 0.06, $labelWidth + 14.5),
            $area->Top + $topReserve, $area->Right - $area->Width * 0.031,
            $area->Bottom - $bottomReserve);
    }
}
