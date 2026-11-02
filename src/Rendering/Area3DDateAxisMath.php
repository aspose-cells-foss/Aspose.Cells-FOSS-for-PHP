<?php

namespace Aspose\Cells_FOSS\Rendering;

final class Area3DDateAxisMath
{
    private function __construct()
    {
    }

    public static function CategoryPosition($chart, $index, $count, $dateSystem)
    {
        if ($chart !== null
            && $chart->HasDateCategoryAxis
            && $index >= 0
            && $index < count($chart->CategoryValues)
            && $chart->CategoryValues[$index] !== null) {
            $minDate = 0.0;
            $maxDate = 0.0;
            if (ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $dateSystem, $minDate, $maxDate)
                && $maxDate > $minDate) {
                $serial = ChartDateAxisMath::NormalizePointSerial($chart, $chart->CategoryValues[$index], $dateSystem);
                $position = ($serial - $minDate) / ($maxDate - $minDate);
                $position = self::Clamp($position);
                return $chart->CategoryAxisMinMax ? $position : 1.0 - $position;
            }
        }

        if ($count <= 1) {
            return 0.5;
        }

        if ($chart !== null && $chart->CrossBetween) {
            return ($index + 0.5) / $count;
        }

        return $index / ($count - 1);
    }

    public static function BuildMonthlyTicks($chart, $dateSystem, $culture)
    {
        $result = array();
        if ($chart === null
            || !$chart->HasDateCategoryAxis
            || $chart->CategoryBaseTimeUnit !== 'months') {
            return $result;
        }

        $minDate = 0.0;
        $maxDate = 0.0;
        $plotMinDate = 0.0;
        $plotMaxDate = 0.0;
        if (!ChartDateAxisMath::TryGetDateAxisRange($chart, $minDate, $maxDate)
            || !ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $dateSystem, $plotMinDate, $plotMaxDate)) {
            return $result;
        }

        $tick = ChartDateAxisMath::StartOfMonth($minDate, $dateSystem);
        $last = ChartDateAxisMath::StartOfMonth($maxDate, $dateSystem);
        $formatCode = ($chart->CategoryFormatCode === null || $chart->CategoryFormatCode === '') ? 'mmm-yy' : $chart->CategoryFormatCode;
        while ($tick <= $last) {
            $position = self::Clamp(($tick - $plotMinDate) / ($plotMaxDate - $plotMinDate));
            if (!$chart->CategoryAxisMinMax) {
                $position = 1.0 - $position;
            }

            $label = ChartXmlParser::FormatLabel($tick, $formatCode, $dateSystem, $culture);
            $result[] = array($position, $label !== null ? rtrim($label) : '');
            $tick = ChartDateAxisMath::AddMonthsSerial($tick, 1, $dateSystem);
        }

        return $result;
    }

    public static function SortByPosition(&$positions, &$frontPoints, &$backPoints)
    {
        for ($i = 1; $i < count($positions); $i++) {
            $position = $positions[$i];
            $front = $frontPoints[$i];
            $back = $backPoints[$i];
            $j = $i - 1;
            while ($j >= 0 && $positions[$j] > $position) {
                $positions[$j + 1] = $positions[$j];
                $frontPoints[$j + 1] = $frontPoints[$j];
                $backPoints[$j + 1] = $backPoints[$j];
                $j--;
            }

            $positions[$j + 1] = $position;
            $frontPoints[$j + 1] = $front;
            $backPoints[$j + 1] = $back;
        }
    }

    private static function Clamp($value)
    {
        if ($value < 0.0) {
            return 0.0;
        }

        if ($value > 1.0) {
            return 1.0;
        }

        return $value;
    }
}
