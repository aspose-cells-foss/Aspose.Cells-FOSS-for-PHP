<?php

namespace Aspose\Cells_FOSS\Rendering;

final class ChartValueAxisIntervals
{
    private function __construct()
    {
    }

    public static function ResolveMajor($chart, $plot, $major)
    {
        if ($chart !== null && $chart->Kind === ChartKind::Scatter && $chart->MajorUnit === null
            && $plot->Height < 105.0) {
            return $major * 2.0;
        }

        if ($chart === null || !$chart->IsPercentStacked || $chart->MajorUnit !== null || $major < 0.1) {
            return $major;
        }

        if ($chart->HasDataTable) {
            return 0.5;
        }

        // Ten intervals need at least 12 pt each for the 9 pt axis labels.
        // Compact 2D line plots use five intervals; taller plots retain ten.
        if ($chart->Kind === ChartKind::Line && !$chart->IsLine3D && $plot->Height < 120.0) {
            return 0.2;
        }

        return 0.1;
    }
}
