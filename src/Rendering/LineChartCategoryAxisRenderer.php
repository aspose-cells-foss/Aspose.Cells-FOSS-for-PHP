<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;

final class LineChartCategoryAxisRenderer
{
    private function __construct()
    {
    }

    public static function DrawBoundaryGridlines($canvas, $plot, $chart, $dates)
    {
        if (!$chart->HasCategoryGridlines || $chart->Kind !== ChartKind::Line || $chart->IsLine3D
            || !$chart->HasDateCategoryAxis || !$chart->CrossBetween
            || $chart->CategoryBaseTimeUnit !== 'months') {
            return false;
        }
        $positions = self::TickPositions($plot, $chart, $dates);
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->CategoryGridlineColor;
        $paint->StrokeWidth = (float)max(0.5, $chart->CategoryGridlineWidthPt);
        $paint->IsAntialias = false;
        for ($i = 0; $i < count($positions); $i++) {
            $canvas->DrawLine($positions[$i], $plot->Top, $positions[$i], $plot->Bottom, $paint);
        }
        $paint->Dispose();
        return true;
    }

    public static function Draw($canvas, $plot, $chart, $dates)
    {
        if (!$chart->HasCategoryAxisLine) {
            return;
        }
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->CategoryAxisLineColor;
        $paint->StrokeWidth = (float)max(0.5, $chart->CategoryAxisLineWidthPt);
        $paint->IsAntialias = false;
        $canvas->DrawLine($plot->Left, $plot->Bottom, $plot->Right, $plot->Bottom, $paint);
        $inside = $chart->CategoryAxisMajorTickMark == 'in' || $chart->CategoryAxisMajorTickMark == 'cross';
        $outside = $chart->CategoryAxisMajorTickMark == 'out' || $chart->CategoryAxisMajorTickMark == 'cross';
        if (!$inside && !$outside) {
            $paint->Dispose();
            return;
        }
        $positions = self::TickPositions($plot, $chart, $dates);
        for ($i = 0; $i < count($positions); $i++) {
            $canvas->DrawLine($positions[$i], $plot->Bottom - ($inside ? 2.8 : 0.0),
                $positions[$i], $plot->Bottom + ($outside ? 2.8 : 0.0), $paint);
        }
        $paint->Dispose();
    }

    public static function TickPositions($plot, $chart, $dates)
    {
        $positions = array();
        if (!$chart->HasDateCategoryAxis) {
            $count = count($chart->Categories);
            $intervals = $chart->CrossBetween ? $count : $count - 1;
            if ($intervals <= 0) {
                return $positions;
            }
            for ($i = 0; $i <= $intervals; $i++) {
                $positions[] = $plot->Left + $plot->Width * $i / $intervals;
            }
            return $positions;
        }
        if ($chart->CategoryBaseTimeUnit !== 'months') {
            return $positions;
        }
        $min = 0.0;
        $max = 0.0;
        $plotMin = 0.0;
        $plotMax = 0.0;
        if (!ChartDateAxisMath::TryGetDateAxisRange($chart, $min, $max)
            || !ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $dates, $plotMin, $plotMax)) {
            return $positions;
        }
        $current = ChartDateAxisMath::StartOfMonth($min, $dates);
        $last = ChartDateAxisMath::StartOfMonth($max, $dates);
        if ($chart->CrossBetween) {
            $last = ChartDateAxisMath::AddMonthsSerial($last, 1, $dates);
        }
        while ($current <= $last + 0.1) {
            $serial = $current;
            if ($chart->CrossBetween) {
                $serial = (ChartDateAxisMath::AddMonthsSerial($current, -1, $dates) + $current) * 0.5;
            }
            $positions[] = ChartDateAxisMath::DateCategoryX($serial, $plotMin, $plotMax, $plot, $chart, $dates);
            $current = ChartDateAxisMath::AddMonthsSerial($current, 1, $dates);
        }
        return $positions;
    }
}
