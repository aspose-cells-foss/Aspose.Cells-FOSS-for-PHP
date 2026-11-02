<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKRect;

final class DoughnutChartRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $area, $chart)
    {
        $size = min($area->Width, $area->Height) * 0.635;
        $rect = new SKRect($area->MidX - $size * 0.5, $area->Top + $area->Height * 0.19,
            $area->MidX + $size * 0.5, $area->Top + $area->Height * 0.19 + $size);
        if ($canvas === null || $chart === null || count($chart->Series) === 0) {
            return $rect;
        }

        $outerRadius = $size * 0.5;
        $holeRadius = $outerRadius * (float)($chart->DoughnutHoleSizePercent / 100.0);
        $ringWidth = ($outerRadius - $holeRadius) / count($chart->Series);
        $centerX = $rect->MidX;
        $centerY = $rect->MidY;

        for ($ringIndex = 0; $ringIndex < count($chart->Series); $ringIndex++) {
            $seriesIndex = count($chart->Series) - $ringIndex - 1;
            $values = $chart->Series[$seriesIndex]->Values;
            $sum = 0.0;
            for ($i = 0; $i < count($values); $i++) {
                if ($values[$i] !== null) {
                    $sum += abs($values[$i]);
                }
            }

            if ($sum <= 0.0) {
                continue;
            }

            $ringOuter = $outerRadius - $ringIndex * $ringWidth;
            $ringInner = $ringOuter - $ringWidth;
            $outerRect = SKRect::Create($centerX - $ringOuter, $centerY - $ringOuter,
                $ringOuter * 2.0, $ringOuter * 2.0);
            $innerRect = SKRect::Create($centerX - $ringInner, $centerY - $ringInner,
                $ringInner * 2.0, $ringInner * 2.0);
            $start = (float)(-90.0 + $chart->DoughnutFirstSliceAngleDeg);
            for ($i = 0; $i < count($values); $i++) {
                if ($values[$i] === null) {
                    continue;
                }

                $sweep = (float)(abs($values[$i]) / $sum * 360.0);
                $startRadians = $start * (float)pi() / 180.0;
                $endRadians = ($start + $sweep) * (float)pi() / 180.0;
                $color = $i < count($chart->PiePointColors) ? $chart->PiePointColors[$i] : $chart->Series[$seriesIndex]->Color;
                $path = new SKPath();
                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = $color;
                $fill->IsAntialias = true;
                $path->MoveTo($centerX + $ringOuter * (float)cos($startRadians),
                    $centerY + $ringOuter * (float)sin($startRadians));
                $path->ArcTo($outerRect, $start, $sweep, false);
                $path->LineTo($centerX + $ringInner * (float)cos($endRadians),
                    $centerY + $ringInner * (float)sin($endRadians));
                $path->ArcTo($innerRect, $start + $sweep, -$sweep, false);
                $path->Close();
                $canvas->DrawPath($path, $fill);
                $fill->Dispose();
                $path->Dispose();

                $start += $sweep;
            }
        }

        return $rect;
    }
}
