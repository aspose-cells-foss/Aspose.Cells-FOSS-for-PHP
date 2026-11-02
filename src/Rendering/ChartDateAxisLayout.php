<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSerialConverter;

final class ChartDateAxisLayout
{
    public $MonthStep;

    public $RotationDeg;

    public static function Resolve($chart, $dateSystem, $availableWidth)
    {
        $result = new ChartDateAxisLayout();
        $result->MonthStep = 1;
        $result->RotationDeg = 0.0;

        if ($chart === null || !$chart->HasDateCategoryAxis || $chart->CategoryValues === null) {
            return $result;
        }

        if (!self::TryGetDateAxisRange($chart, $minDate, $maxDate)) {
            return $result;
        }

        $visibleCategoryCount = self::CountVisibleCategoryValues($chart);
        $monthCount = self::GetInclusiveMonthCount($minDate, $maxDate, $dateSystem);
        if (abs($chart->CategoryAxisTextRotationDeg) <= 0.1) {
            return $result;
        }

        $visibleTickCount = self::CountVisibleTickCount($monthCount, $result->MonthStep);
        $slotWidth = $availableWidth;
        if ($visibleTickCount > 1) {
            $slotWidth = $availableWidth / ($visibleTickCount - 1);
        }

        if (LineChartAutomaticLayout::Applies($chart)) {
            // Choose orientation from available tick spacing, not a literal angle
            // obtained by clamping Excel's automatic-orientation marker.
            $result->RotationDeg = $slotWidth >= 52.0 ? 0.0 : ($slotWidth >= 18.0 ? -45.0 : -90.0);
        } elseif ($monthCount > $visibleCategoryCount + 2 || $slotWidth < 52.0) {
            $result->RotationDeg = (float)$chart->CategoryAxisTextRotationDeg;
        }

        return $result;
    }

    private static function TryGetDateAxisRange($chart, &$minDate, &$maxDate)
    {
        $minDate = PHP_FLOAT_MAX;
        $maxDate = -PHP_FLOAT_MAX;
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

    private static function CountVisibleCategoryValues($chart)
    {
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

    private static function CountVisibleTickCount($monthCount, $monthStep)
    {
        if ($monthCount <= 0) {
            return 0;
        }

        if ($monthStep <= 1) {
            return $monthCount;
        }

        return intdiv($monthCount - 1, $monthStep) + 1;
    }
}
