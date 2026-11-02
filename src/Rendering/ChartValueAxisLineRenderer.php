<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;

final class ChartValueAxisLineRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $plot, $chart, $min, $max, $major)
    {
        if (!$chart->HasValueAxisLine || $max <= $min || $major <= 0.0) {
            return;
        }

        $axisX = $chart->ValueAxisOnRight ? $plot->Right : $plot->Left;
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->ValueAxisLineColor;
        $paint->StrokeWidth = (float)max(0.5, $chart->ValueAxisLineWidthPt);
        $paint->IsAntialias = false;
        $canvas->DrawLine($axisX, $plot->Bottom, $axisX, $plot->Top, $paint);

        $tickMark = $chart->ValueAxisMajorTickMark;
        if ($tickMark !== 'cross' && $tickMark !== 'in' && $tickMark !== 'out') {
            $paint->Dispose();
            return;
        }

        $inward = $chart->ValueAxisOnRight ? -2.8 : 2.8;
        $start = $tickMark === 'cross' || $tickMark === 'out' ? -$inward : 0.0;
        $end = $tickMark === 'cross' || $tickMark === 'in' ? $inward : 0.0;
        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $y = $plot->Bottom - (float)(($value - $min) / ($max - $min)) * $plot->Height;
            $canvas->DrawLine($axisX + $start, $y, $axisX + $end, $y, $paint);
        }

        $paint->Dispose();
    }
}
