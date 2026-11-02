<?php

namespace Aspose\Cells_FOSS\Rendering;

final class ScatterChartAxisMath
{
    private function __construct()
    {
    }

    public static function ResolveXRange($chart, &$minimum, &$maximum, &$major)
    {
        $dataMinimum = PHP_FLOAT_MAX;
        $dataMaximum = -PHP_FLOAT_MAX;
        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $values = $chart->Series[$seriesIndex]->XValues;
            for ($pointIndex = 0; $pointIndex < count($values); $pointIndex++) {
                if ($values[$pointIndex] === null) {
                    continue;
                }

                $value = $values[$pointIndex];
                $dataMinimum = min($dataMinimum, $value);
                $dataMaximum = max($dataMaximum, $value);
            }
        }

        if ($dataMinimum === PHP_FLOAT_MAX) {
            $minimum = 0.0;
            $maximum = 1.0;
            $major = 0.2;
            return;
        }

        if ($chart->IsBubble) {
            BubbleChartMath::ResolveXRange($chart, $dataMinimum, $dataMaximum,
                $minimum, $maximum, $major);
            return;
        }

        $isDateAxis = ChartXmlParser::IsDateFormat($chart->ScatterXFormatCode);
        $automaticMajor = self::NiceNumber(($dataMaximum - $dataMinimum) / 8.0);
        $automaticMinimum = !$isDateAxis && $dataMinimum >= 0.0
            ? 0.0
            : floor($dataMinimum / $automaticMajor) * $automaticMajor;
        $minimum = $chart->ScatterXMin !== null ? $chart->ScatterXMin : $automaticMinimum;
        $major = $chart->ScatterXMajorUnit !== null
            ? $chart->ScatterXMajorUnit
            : ($isDateAxis
                ? $automaticMajor
                : self::NiceNumber(($dataMaximum - $minimum) / 8.0));
        if ($major <= 0.0) {
            $major = 1.0;
        }

        $maximum = $chart->ScatterXMax !== null
            ? $chart->ScatterXMax
            : ceil($dataMaximum / $major) * $major;
        if ($chart->ScatterXMax === null && ($maximum - $dataMaximum) / $major < 0.2) {
            $maximum += $major;
        }
        if ($isDateAxis && $chart->ScatterXMin === null && ($dataMinimum - $minimum) / $major < 0.2) {
            $minimum -= $major;
        }
        if ($maximum <= $minimum) {
            $maximum = $minimum + $major;
        }
    }

    public static function ToScreen($value, $minimum, $maximum, $left, $right)
    {
        if ($maximum <= $minimum) {
            return $left;
        }

        return $left + (float)(($value - $minimum) / ($maximum - $minimum)) * ($right - $left);
    }

    private static function NiceNumber($value)
    {
        if ($value <= 0.0) {
            return 1.0;
        }

        $exponent = floor(log10($value));
        $fraction = $value / pow(10.0, $exponent);
        if ($fraction < 1.5) {
            $nice = 1.0;
        } elseif ($fraction < 3.0) {
            $nice = 2.0;
        } elseif ($fraction < 7.0) {
            $nice = 5.0;
        } else {
            $nice = 10.0;
        }
        return $nice * pow(10.0, $exponent);
    }
}
