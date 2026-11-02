<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSerialConverter;

final class ChartDateAxisMath
{
    private function __construct()
    {
    }

    public static function TryGetDateAxisPlotRange($chart, $dateSystem, &$plotMinDate, &$plotMaxDate)
    {
        $plotMinDate = 0.0;
        $plotMaxDate = 0.0;

        if (!self::TryGetDateAxisRange($chart, $minDate, $maxDate)) {
            return false;
        }

        if ($chart === null || !$chart->CrossBetween) {
            $plotMinDate = $minDate;
            $plotMaxDate = $maxDate;
            return true;
        }

        if ($chart->CategoryBaseTimeUnit === 'months') {
            $firstTick = self::StartOfMonth($minDate, $dateSystem);
            $lastTick = self::StartOfMonth($maxDate, $dateSystem);
            $previousTick = self::AddMonthsSerial($firstTick, -1, $dateSystem);
            $nextTick = self::AddMonthsSerial($lastTick, 1, $dateSystem);
            $plotMinDate = self::Midpoint($previousTick, $firstTick);
            $plotMaxDate = self::Midpoint($lastTick, $nextTick);
            return $plotMaxDate > $plotMinDate;
        }

        if ($chart->CategoryValues !== null && count($chart->CategoryValues) > 1) {
            if (self::TryGetFirstTwoCategoryValues($chart, $firstValue, $secondValue)
                && self::TryGetLastTwoCategoryValues($chart, $previousValue, $lastValue)) {
                $plotMinDate = $firstValue - ($secondValue - $firstValue) * 0.5;
                $plotMaxDate = $lastValue + ($lastValue - $previousValue) * 0.5;
                return $plotMaxDate > $plotMinDate;
            }
        }

        $plotMinDate = $minDate;
        $plotMaxDate = $maxDate;
        return true;
    }

    public static function TryGetDateAxisRange($chart, &$minDate, &$maxDate)
    {
        $minDate = PHP_FLOAT_MAX;
        $maxDate = -PHP_FLOAT_MAX;
        if ($chart === null || $chart->CategoryValues === null) {
            return false;
        }

        for ($i = 0; $i < count($chart->CategoryValues); $i++) {
            if ($chart->CategoryValues[$i] === null) {
                continue;
            }

            $value = $chart->CategoryValues[$i];
            if ($value < $minDate) {
                $minDate = $value;
            }

            if ($value > $maxDate) {
                $maxDate = $value;
            }
        }

        if ($minDate === PHP_FLOAT_MAX || $maxDate === -PHP_FLOAT_MAX) {
            $minDate = 0.0;
            $maxDate = 0.0;
            return false;
        }

        return true;
    }

    private static function TryGetFirstTwoCategoryValues($chart, &$firstValue, &$secondValue)
    {
        $firstValue = 0.0;
        $secondValue = 0.0;
        $found = 0;
        for ($i = 0; $i < count($chart->CategoryValues); $i++) {
            if ($chart->CategoryValues[$i] === null) {
                continue;
            }

            if ($found === 0) {
                $firstValue = $chart->CategoryValues[$i];
                $found = 1;
            } else {
                $secondValue = $chart->CategoryValues[$i];
                return true;
            }
        }

        return false;
    }

    private static function TryGetLastTwoCategoryValues($chart, &$previousValue, &$lastValue)
    {
        $previousValue = 0.0;
        $lastValue = 0.0;
        $foundLast = false;
        for ($i = count($chart->CategoryValues) - 1; $i >= 0; $i--) {
            if ($chart->CategoryValues[$i] === null) {
                continue;
            }

            if (!$foundLast) {
                $lastValue = $chart->CategoryValues[$i];
                $foundLast = true;
            } else {
                $previousValue = $chart->CategoryValues[$i];
                return true;
            }
        }

        return false;
    }

    private static function Midpoint($a, $b)
    {
        return $a + ($b - $a) * 0.5;
    }

    public static function StartOfMonth($serial, $dateSystem)
    {
        $date = DateSerialConverter::FromSerial($serial, $dateSystem);
        $first = new \DateTime($date->format('Y-m') . '-01 00:00:00');
        return DateSerialConverter::ToSerial($first, $dateSystem);
    }

    public static function AddMonthsSerial($serial, $months, $dateSystem)
    {
        $date = DateSerialConverter::FromSerial($serial, $dateSystem);
        $next = new \DateTime($date->format('Y-m') . '-01 00:00:00');
        $next->modify(($months >= 0 ? '+' : '-') . abs($months) . ' month');
        return DateSerialConverter::ToSerial($next, $dateSystem);
    }

    public static function DateCategoryX($serial, $minDate, $maxDate, $plot, $chart = null, $dateSystem = null)
    {
        if ($chart !== null
            && LineChartAutomaticLayout::Applies($chart)
            && $chart->CategoryBaseTimeUnit === 'months') {
            // A monthly date axis has equal-width months, including February.
            return self::DateCategoryX(self::MonthCoordinate($serial, $dateSystem), self::MonthCoordinate($minDate, $dateSystem),
                self::MonthCoordinate($maxDate, $dateSystem), $plot);
        }

        if ($maxDate <= $minDate) {
            return $plot->MidX;
        }

        $t = ($serial - $minDate) / ($maxDate - $minDate);
        if ($t < 0.0) {
            $t = 0.0;
        } elseif ($t > 1.0) {
            $t = 1.0;
        }

        return $plot->Left + (float)$t * $plot->Width;
    }

    private static function MonthCoordinate($serial, $dateSystem)
    {
        $date = DateSerialConverter::FromSerial($serial, $dateSystem);
        $daysInMonth = (int)$date->format('t');
        $timeOfDayDays = ((int)$date->format('G') * 3600 + (int)$date->format('i') * 60
            + (int)$date->format('s') + (int)$date->format('u') / 1000000.0) / 86400.0;
        return (int)$date->format('Y') * 12.0 + (int)$date->format('n') - 1.0
            + ((int)$date->format('j') - 1.0 + $timeOfDayDays) / $daysInMonth;
    }

    public static function NormalizePointSerial($chart, $serial, $dateSystem)
    {
        if ($chart === null || !$chart->HasDateCategoryAxis) {
            return $serial;
        }

        if ($chart->CategoryBaseTimeUnit !== 'months') {
            return $serial;
        }

        if (!self::TryGetDateAxisRange($chart, $minDate, $maxDate)) {
            return $serial;
        }

        $visibleCategoryCount = self::CountVisibleCategoryValues($chart);
        $monthCount = self::GetInclusiveMonthCount($minDate, $maxDate, $dateSystem);
        if ($visibleCategoryCount > 0 && ($monthCount <= $visibleCategoryCount + 2
            || LineChartAutomaticLayout::Applies($chart) || $chart->IsLine3D)) {
            return self::StartOfMonth($serial, $dateSystem);
        }

        return $serial;
    }

    private static function CountVisibleCategoryValues($chart)
    {
        if ($chart === null || $chart->CategoryValues === null) {
            return 0;
        }

        $count = 0;
        for ($i = 0; $i < count($chart->CategoryValues); $i++) {
            if ($chart->CategoryValues[$i] !== null) {
                $count++;
            }
        }

        return $count;
    }

    private static function GetInclusiveMonthCount($minDate, $maxDate, $dateSystem)
    {
        $min = DateSerialConverter::FromSerial($minDate, $dateSystem);
        $max = DateSerialConverter::FromSerial($maxDate, $dateSystem);
        return ((int)$max->format('Y') - (int)$min->format('Y')) * 12 + (int)$max->format('n') - (int)$min->format('n') + 1;
    }
}
