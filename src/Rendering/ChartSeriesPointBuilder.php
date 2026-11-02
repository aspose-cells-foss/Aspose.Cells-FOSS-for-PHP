<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPoint;

final class ChartSeriesPointBuilder
{
    private function __construct()
    {
    }

    public static function BuildDateSeriesPoints($chart, $values, $plot, $min, $max, $dateSystem)
    {
        if ($chart === null || !$chart->HasDateCategoryAxis || $chart->CategoryValues === null || $values === null) {
            return null;
        }

        if (!ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $dateSystem, $plotMinDate, $plotMaxDate)) {
            return null;
        }

        $serials = array();
        $points = array();
        $count = count($chart->CategoryValues);
        if (count($values) < $count) {
            $count = count($values);
        }

        for ($i = 0; $i < $count; $i++) {
            if ($chart->CategoryValues[$i] === null || $values[$i] === null) {
                continue;
            }

            $serial = ChartDateAxisMath::NormalizePointSerial($chart, $chart->CategoryValues[$i], $dateSystem);
            $serials[] = $serial;
            $points[] = new SKPoint(
                ChartDateAxisMath::DateCategoryX($serial, $plotMinDate, $plotMaxDate, $plot, $chart, $dateSystem),
                self::ValueToY($values[$i], $min, $max, $plot));
        }

        for ($i = 1; $i < count($serials); $i++) {
            $serial = $serials[$i];
            $point = $points[$i];
            $j = $i - 1;
            while ($j >= 0 && $serials[$j] > $serial) {
                $serials[$j + 1] = $serials[$j];
                $points[$j + 1] = $points[$j];
                $j--;
            }

            $serials[$j + 1] = $serial;
            $points[$j + 1] = $point;
        }

        return $points;
    }

    private static function ValueToY($value, $min, $max, $plot)
    {
        $t = ($value - $min) / ($max - $min);
        return $plot->Bottom - (float)$t * $plot->Height;
    }
}
