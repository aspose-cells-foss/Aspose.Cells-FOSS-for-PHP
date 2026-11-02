<?php

namespace Aspose\Cells_FOSS\Rendering;

final class LineChartDecorationLayout
{
    private function __construct()
    {
    }

    private static function UsesManualLinePlot($chart)
    {
        return $chart !== null && $chart->Kind === ChartKind::Line
            && !$chart->IsLine3D && $chart->HasManualPlot;
    }

    public static function TitleCenterY($area, $chart, $fallback)
    {
        if ($chart !== null && $chart->IsLine3D && !$chart->HasDateCategoryAxis && !$chart->HasManualTitle) {
            return $area->Top + ($chart->TitleFont !== null ? (float)$chart->TitleFont->Size : 14.0) * 1.05;
        }
        if ((!self::UsesManualLinePlot($chart) && !LineChartAutomaticLayout::Applies($chart)
            && !LineChartAutomaticLayout::AppliesToCategoryAxis($chart)) || $chart->HasManualTitle) {
            return $fallback;
        }
        // Keep the automatic title in its chart band, independent of the inner plot.
        $size = $chart->TitleFont !== null ? (float)$chart->TitleFont->Size : 14.0;
        return $area->Top + $size * 1.05;
    }

    public static function LegendCenterY($area, $chart, $fallback)
    {
        if ($chart !== null && $chart->IsLine3D && !$chart->HasDateCategoryAxis && !$chart->HasManualLegend
            && $chart->LegendPosition === 'b') {
            return $area->Bottom - 15.7;
        }
        if (!self::UsesAutomaticBottomLegend($chart)) {
            return $fallback;
        }
        // Anchor the automatic legend to the chart, not the remaining plot gap.
        return $area->Bottom - ($chart->HasManualPlot ? 15.2 : 14.8);
    }

    public static function UsesAutomaticBottomLegend($chart)
    {
        return (self::UsesManualLinePlot($chart) || LineChartAutomaticLayout::Applies($chart)
            || LineChartAutomaticLayout::AppliesToCategoryAxis($chart)) && !$chart->HasManualLegend
            && $chart->LegendPosition === 'b';
    }

    public static function SwatchWidth($chart, $fallback)
    {
        return self::UsesAutomaticBottomLegend($chart) ? 19.2 : $fallback;
    }

    public static function TextGap($chart, $fallback)
    {
        return self::UsesAutomaticBottomLegend($chart) ? 2.0 : $fallback;
    }

    public static function ItemGap($chart, $fallback)
    {
        if (self::UsesAutomaticBottomLegend($chart) && LineChartAutomaticLayout::AppliesToCategoryAxis($chart)) {
            return 8.5;
        }
        return self::UsesAutomaticBottomLegend($chart) ? 9.5 : $fallback;
    }

    public static function HorizontalOffset($chart, $fallback)
    {
        // Match the optical centering of the line-key and text group.
        if (self::UsesAutomaticBottomLegend($chart) && LineChartAutomaticLayout::AppliesToCategoryAxis($chart)) {
            return 2.25;
        }
        return self::UsesAutomaticBottomLegend($chart) && !$chart->HasManualPlot
            && LineChartAutomaticLayout::Applies($chart)
            ? 2.1 : $fallback;
    }
}
