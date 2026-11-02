<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;

/**
 * Draws the projected top and visible front rim of a three-dimensional pie.
 */
final class Pie3DChartRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $area, $chart, $values, $colors, $sum)
    {
        $perspective = $chart->RightAngleAxes ? 0.0
            : (float)(sin($chart->PerspectiveDeg * M_PI / 180.0) * 0.68);
        $perspective = max(0.0, min(0.6, $perspective));
        $flatView = 1.0 - min(1.0, $perspective / 0.34);
        $projectedRadiusX = $area->Width * (0.45 - 0.03 * $flatView);
        $radiusX = $projectedRadiusX * (float)sqrt(1.0 - $perspective * $perspective);
        $tilt = (float)sin(max(0.0, min(70.0, $chart->RotationXDeg)) * M_PI / 180.0);
        $radiusY = $area->Height * (0.15 + 0.065 * $flatView) * max(0.2,
            $tilt / (float)sin(15.0 * M_PI / 180.0));
        $centerY = $area->Top + $area->Height * (0.303 + 0.107 * $flatView);
        $depth = $area->Height * 0.193 * (float)max(0.25,
            min(2.0, $chart->DepthPercent / 100.0));
        $center = new SKPoint($area->MidX, $centerY);
        $ellipse = new SKRect($center->X - $projectedRadiusX,
            $center->Y - $radiusY / (1.0 + $perspective),
            $center->X + $projectedRadiusX,
            $center->Y + $radiusY / (1.0 - $perspective));

        $start = -90.0 + (float)$chart->RotationYDeg;
        $sliceStarts = array_fill(0, count($values), 0.0);
        for ($i = 0; $i < count($values); $i++) {
            $sliceStarts[$i] = $start;
            if ($values[$i] === null) {
                continue;
            }

            $sweep = (float)(abs($values[$i]) / $sum * 360.0);
            self::DrawVisibleRim($canvas, $center, $radiusX, $radiusY, $perspective,
                $depth, $start, $sweep, $colors[$i]);
            $start += $sweep;
        }

        for ($i = count($values) - 1; $i >= 0; $i--) {
            if ($values[$i] === null) {
                continue;
            }

            $sweep = (float)(abs($values[$i]) / $sum * 360.0);
            $path = new SKPath();
            $fill = new SKPaint();
            $fill->Style = SKPaintStyle::Fill;
            $fill->Color = self::Darken($colors[$i], 0.92);
            $fill->IsAntialias = true;
            $path->MoveTo($center);
            self::AddProjectedArc($path, $center, $radiusX, $radiusY, $perspective,
                $depth, $sliceStarts[$i], $sliceStarts[$i] + $sweep, false, false);
            $path->Close();
            $canvas->DrawPath($path, $fill);
            $fill->Dispose();
            $path->Dispose();
        }

        return $ellipse;
    }

    private static function DrawVisibleRim($canvas, $center, $radiusX, $radiusY, $perspective, $depth, $start, $sweep, $color)
    {
        $silhouetteAngle = (float)(asin($perspective) * 180.0 / M_PI);
        $visibleStart = max($silhouetteAngle, $start);
        $visibleEnd = min(180.0 - $silhouetteAngle, $start + $sweep);
        if ($visibleEnd <= $visibleStart) {
            return;
        }

        $path = new SKPath();
        self::AddProjectedArc($path, $center, $radiusX, $radiusY, $perspective,
            $depth, $visibleStart, $visibleEnd, false, true);
        self::AddProjectedArc($path, $center, $radiusX, $radiusY, $perspective,
            $depth, $visibleEnd, $visibleStart, true, false);
        $path->Close();

        $leftShade = self::Darken($color, 0.95);
        $centerShade = self::Darken($color, 0.70);
        $rightShade = self::Darken($color, 0.35);
        $shader = SKShader::CreateLinearGradient(
            new SKPoint($center->X - $radiusX, 0.0),
            new SKPoint($center->X + $radiusX, 0.0),
            array($leftShade, $centerShade, $rightShade),
            array(0.0, 0.5, 1.0), SKShaderTileMode::Clamp);
        $fill = new SKPaint();
        $fill->Style = SKPaintStyle::Fill;
        $fill->Shader = $shader;
        $fill->IsAntialias = true;
        $canvas->DrawPath($path, $fill);
        $fill->Dispose();
        $shader->Dispose();
        $path->Dispose();
    }

    private static function AddProjectedArc($path, $center, $radiusX, $radiusY, $perspective, $depth, $start, $end, $lower, $moveToStart)
    {
        $segments = max(1, (int)ceil(abs($end - $start) / 2.0));
        for ($i = 0; $i <= $segments; $i++) {
            $angle = $start + ($end - $start) * $i / $segments;
            $radians = $angle * (float)M_PI / 180.0;
            $sine = (float)sin($radians);
            $divisor = 1.0 - $perspective * $sine;
            $x = $center->X + $radiusX * (float)cos($radians) / $divisor;
            $y = $center->Y + ($radiusY * $sine + ($lower ? $depth : 0.0)) / $divisor;
            if ($lower) {
                $x = $center->X + ($x - $center->X) * 0.978;
            }

            if ($i === 0 && $moveToStart) {
                $path->MoveTo($x, $y);
            } else {
                $path->LineTo($x, $y);
            }
        }
    }

    private static function Darken($color, $factor)
    {
        return new SKColor((int)($color->Red * $factor),
            (int)($color->Green * $factor), (int)($color->Blue * $factor), $color->Alpha);
    }
}
