<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKRect;

final class ScatterChartPlotLayout
{
    private function __construct()
    {
    }

    public static function Resolve($area, $chart)
    {
        $dateAxis = ChartXmlParser::IsDateFormat($chart->ScatterXFormatCode);
        $rightReserve = $chart->LegendPosition === null ? 0.05
            : ($dateAxis ? 0.06 : 0.078);
        $leftReserve = $chart->LegendPosition === null ? 0.10
            : ($chart->ValueDisplayUnitLabel === null || $chart->ValueDisplayUnitLabel === '' ? 0.132 : 0.162);
        $bottomReserve = $chart->LegendPosition === null ? 0.11 : 0.235;
        if ($chart->IsBubble && $chart->LegendPosition !== null) {
            $bottomReserve = 0.173;
        }
        if ($chart->ValueAxisTitle !== null && $chart->ValueAxisTitle !== '') {
            $leftReserve += 0.057;
        }
        if ($chart->CategoryAxisTitle !== null && $chart->CategoryAxisTitle !== '') {
            $bottomReserve += 0.116;
        }

        // Tick, title and legend bands stop growing when a chart spans extra rows or columns.
        $bandWidth = min($area->Width, 384.4);
        $bandHeight = min($area->Height, 195.4);
        $rightBandWidth = $dateAxis ? $area->Width : min($area->Width, 400.0);
        return new SKRect($area->Left + $bandWidth * $leftReserve,
            $area->Top + $bandHeight * 0.187,
            $area->Right - $rightBandWidth * $rightReserve,
            $area->Bottom - $bandHeight * $bottomReserve);
    }
}
