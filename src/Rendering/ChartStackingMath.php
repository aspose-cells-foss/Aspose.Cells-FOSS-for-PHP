<?php

namespace Aspose\Cells_FOSS\Rendering;

final class ChartStackingMath
{
    private function __construct()
    {
    }

    public static function ComputeDataRange($chart, &$minValue, &$maxValue)
    {
        $minValue = PHP_FLOAT_MAX;
        $maxValue = -PHP_FLOAT_MAX;

        if ($chart === null) {
            return;
        }

        if (!$chart->IsStacked && !$chart->IsPercentStacked) {
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                for ($i = 0; $i < count($series->Values); $i++) {
                    if ($series->Values[$i] === null) {
                        continue;
                    }

                    $value = $series->Values[$i];
                    if ($value < $minValue) {
                        $minValue = $value;
                    }

                    if ($value > $maxValue) {
                        $maxValue = $value;
                    }
                }
            }

            return;
        }

        if ($chart->IsPercentStacked) {
            $minValue = 0.0;
            $maxValue = 1.0;
            return;
        }

        $pointCount = 0;
        for ($s = 0; $s < count($chart->Series); $s++) {
            if (count($chart->Series[$s]->Values) > $pointCount) {
                $pointCount = count($chart->Series[$s]->Values);
            }
        }

        for ($i = 0; $i < $pointCount; $i++) {
            $positive = 0.0;
            $negative = 0.0;
            $sawValue = false;

            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $sawValue = true;
                $value = $series->Values[$i];
                if ($value >= 0.0) {
                    $positive += $value;
                } else {
                    $negative += $value;
                }
            }

            if (!$sawValue) {
                continue;
            }

            if ($negative < $minValue) {
                $minValue = $negative;
            }

            if ($positive > $maxValue) {
                $maxValue = $positive;
            }
        }
    }

    public static function BuildDisplayValues($chart, $seriesIndex)
    {
        $result = array();
        if ($chart === null || $seriesIndex < 0 || $seriesIndex >= count($chart->Series)) {
            return $result;
        }

        $targetSeries = $chart->Series[$seriesIndex];
        $pointCount = count($targetSeries->Values);
        if (!$chart->IsStacked && !$chart->IsPercentStacked) {
            for ($i = 0; $i < count($targetSeries->Values); $i++) {
                $result[] = $targetSeries->Values[$i];
            }

            return $result;
        }

        if ($chart->IsPercentStacked) {
            return self::BuildPercentStackedDisplayValues($chart, $seriesIndex, $pointCount);
        }

        for ($i = 0; $i < $pointCount; $i++) {
            $result[] = null;
        }

        $positive = array_fill(0, $pointCount, 0.0);
        $negative = array_fill(0, $pointCount, 0.0);

        for ($s = 0; $s <= $seriesIndex; $s++) {
            $series = $chart->Series[$s];
            $count = min($pointCount, count($series->Values));
            for ($i = 0; $i < $count; $i++) {
                if ($series->Values[$i] === null) {
                    continue;
                }

                $value = $series->Values[$i];
                if ($value >= 0.0) {
                    $positive[$i] += $value;
                    if ($s === $seriesIndex) {
                        $result[$i] = $positive[$i];
                    }
                } else {
                    $negative[$i] += $value;
                    if ($s === $seriesIndex) {
                        $result[$i] = $negative[$i];
                    }
                }
            }
        }

        return $result;
    }

    private static function BuildPercentStackedDisplayValues($chart, $seriesIndex, $pointCount)
    {
        $result = array();
        for ($i = 0; $i < $pointCount; $i++) {
            $result[] = null;
        }

        $positiveTotals = array_fill(0, $pointCount, 0.0);
        $negativeTotals = array_fill(0, $pointCount, 0.0);
        for ($s = 0; $s < count($chart->Series); $s++) {
            $series = $chart->Series[$s];
            $count = min($pointCount, count($series->Values));
            for ($i = 0; $i < $count; $i++) {
                if ($series->Values[$i] === null) {
                    continue;
                }

                $value = $series->Values[$i];
                if ($value >= 0.0) {
                    $positiveTotals[$i] += $value;
                } else {
                    $negativeTotals[$i] += abs($value);
                }
            }
        }

        $cumulativePositive = array_fill(0, $pointCount, 0.0);
        $cumulativeNegative = array_fill(0, $pointCount, 0.0);
        for ($s = 0; $s <= $seriesIndex; $s++) {
            $series = $chart->Series[$s];
            $count = min($pointCount, count($series->Values));
            for ($i = 0; $i < $count; $i++) {
                if ($series->Values[$i] === null) {
                    continue;
                }

                $value = $series->Values[$i];
                if ($value >= 0.0) {
                    $cumulativePositive[$i] += $value;
                    if ($s === $seriesIndex) {
                        $result[$i] = $positiveTotals[$i] > 0.0 ? $cumulativePositive[$i] / $positiveTotals[$i] : 0.0;
                    }
                } else {
                    $cumulativeNegative[$i] += abs($value);
                    if ($s === $seriesIndex) {
                        $result[$i] = $negativeTotals[$i] > 0.0 ? -($cumulativeNegative[$i] / $negativeTotals[$i]) : 0.0;
                    }
                }
            }
        }

        return $result;
    }
}
