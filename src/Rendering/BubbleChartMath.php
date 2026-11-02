<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPoint;

final class BubbleChartMath
{
    private function __construct()
    {
    }

    public static function ResolveXRange($chart, $dataMinimum, $dataMaximum,
        &$minimum, &$maximum, &$major)
    {
        if ($dataMinimum >= 0.0 && !ChartXmlParser::IsDateFormat($chart->ScatterXFormatCode)
            && $chart->ScatterXMin === null && $chart->ScatterXMax === null) {
            $major = $chart->ScatterXMajorUnit !== null
                ? $chart->ScatterXMajorUnit
                : self::NiceNumber($dataMaximum / 8.0);
            if ($major <= 0.0) {
                $major = 1.0;
            }
            $minimum = 0.0;
            $maximum = max(ceil($dataMaximum / $major), 8.0) * $major;
            return;
        }

        self::ResolveRange($dataMinimum, $dataMaximum, $chart->ScatterXMin, $chart->ScatterXMax,
            $chart->ScatterXMajorUnit, $minimum, $maximum, $major);
    }

    public static function ResolveYRange($chart, $dataMinimum, $dataMaximum,
        &$minimum, &$maximum, &$major)
    {
        self::ResolveRange($dataMinimum, $dataMaximum, $chart->AxisMin, $chart->AxisMax,
            $chart->MajorUnit, $minimum, $maximum, $major);
    }

    public static function DrawSeries($canvas, $plot, $chart, $minimumY, $maximumY)
    {
        $minimumX = 0.0;
        $maximumX = 0.0;
        $majorX = 0.0;
        ScatterChartAxisMath::ResolveXRange($chart, $minimumX, $maximumX, $majorX);
        if ($maximumX <= $minimumX || $maximumY <= $minimumY) {
            return;
        }

        $maximumSize = 0.0;
        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $series = $chart->Series[$seriesIndex];
            for ($pointIndex = 0; $pointIndex < count($series->BubbleSizes); $pointIndex++) {
                if ($series->BubbleSizes[$pointIndex] !== null) {
                    $size = $series->BubbleSizes[$pointIndex];
                    if ($size > 0.0 || $chart->ShowNegativeBubbles) {
                        $maximumSize = max($maximumSize, abs($size));
                    }
                }
            }
        }
        if ($maximumSize <= 0.0) {
            return;
        }

        $maximumRadius = $plot->Height * 0.1264 * (float)($chart->BubbleScalePercent / 100.0);
        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $series = $chart->Series[$seriesIndex];
            $count = min(min(count($series->XValues), count($series->Values)),
                count($series->BubbleSizes));
            $fill = new SKPaint();
            $fill->Style = SKPaintStyle::Fill;
            $fill->Color = $series->Color;
            $fill->IsAntialias = true;
            $texture = $series->Bubble3D ? BubbleSphereRenderer::CreateTexture($series->Color) : null;
            for ($pointIndex = 0; $pointIndex < $count; $pointIndex++) {
                if ($series->XValues[$pointIndex] === null || $series->Values[$pointIndex] === null
                    || $series->BubbleSizes[$pointIndex] === null) {
                    continue;
                }

                $xValue = $series->XValues[$pointIndex];
                $yValue = $series->Values[$pointIndex];
                $size = $series->BubbleSizes[$pointIndex];
                if ($xValue < $minimumX || $xValue > $maximumX || $yValue < $minimumY || $yValue > $maximumY
                    || $size == 0.0 || $size < 0.0 && !$chart->ShowNegativeBubbles) {
                    continue;
                }

                $x = ScatterChartAxisMath::ToScreen($xValue, $minimumX, $maximumX, $plot->Left, $plot->Right);
                $y = $plot->Bottom - (float)(($yValue - $minimumY) / ($maximumY - $minimumY)) * $plot->Height;
                $radius = $maximumRadius * (float)sqrt(abs($size) / $maximumSize);
                if ($series->Bubble3D) {
                    BubbleSphereRenderer::DrawImage($canvas, new SKPoint($x, $y), $radius, $texture);
                } else {
                    $canvas->DrawCircle($x, $y, $radius, $fill);
                }
            }
            if ($texture !== null) {
                $texture->Dispose();
            }
            $fill->Dispose();
        }
    }

    private static function ResolveRange($dataMinimum, $dataMaximum,
        $explicitMinimum, $explicitMaximum, $explicitMajor,
        &$minimum, &$maximum, &$major)
    {
        $span = max(0.0, $dataMaximum - $dataMinimum);
        $major = $explicitMajor !== null ? $explicitMajor : self::NiceNumber($span / 6.0);
        if ($major <= 0.0) {
            $major = 1.0;
        }

        $padding = $span > 0.0 ? $span * 0.25 : $major;
        $minimum = $explicitMinimum !== null
            ? $explicitMinimum
            : floor(($dataMinimum - $padding) / $major) * $major;
        $maximum = $explicitMaximum !== null
            ? $explicitMaximum
            : ceil(($dataMaximum + $padding) / $major) * $major;
        if ($maximum <= $minimum) {
            $maximum = $minimum + $major;
        }
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
