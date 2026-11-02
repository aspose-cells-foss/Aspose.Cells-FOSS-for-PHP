<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKAutoCanvasRestore;

final class ChartSeriesStrokeRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $plot, $chart, $path, $paint)
    {
        $restore = new SKAutoCanvasRestore($canvas);
        if ($chart->Kind === ChartKind::Line && !$chart->IsLine3D) {
            // A line on the axis maximum keeps its center at the data value,
            // but the outer half of its stroke must not escape the plot.
            $canvas->ClipRect($plot);
        }

        $canvas->DrawPath($path, $paint);
        $restore->Dispose();
    }
}
