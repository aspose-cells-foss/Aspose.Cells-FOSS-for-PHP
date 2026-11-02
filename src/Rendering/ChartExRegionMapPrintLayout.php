<?php

namespace Aspose\Cells_FOSS\Rendering;

final class ChartExRegionMapPrintLayout
{
    const Scale = 0.95;

    const LeftInsetPt = 2.5;

    const HorizontalBreakAllowancePt = 10.0;

    private function __construct()
    {
    }

    public static function Applies($sheet)
    {
        if ($sheet === null) {
            return false;
        }

        for ($i = 0; $i < count($sheet->Charts); $i++) {
            $chart = $sheet->Charts[$i];
            if ($chart !== null && $chart->IsChartEx && $chart->RawChartXml !== null
                && strpos((string)$chart->RawChartXml, 'regionMap') !== false) {
                return true;
            }
        }

        return false;
    }

    public static function TopOrigin($pageHeightPt, $topMarginPt)
    {
        // The Excel map-chart print reference places the sheet about 100pt from the top of A4.
        return max($topMarginPt, $pageHeightPt * (100.0 / 842.0));
    }
}
