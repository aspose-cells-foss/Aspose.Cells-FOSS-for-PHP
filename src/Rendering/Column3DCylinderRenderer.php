<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;

final class Column3DCylinderRenderer
{
    const ArcSteps = 16;

    private function __construct()
    {
    }

    public static function ResolveBottomCenterOffset($floorMidpoint, $depth)
    {
        return new SKPoint($floorMidpoint->X - $depth->X * 0.5,
            $floorMidpoint->Y - $depth->Y * 0.5);
    }

    public static function ResolveHeightExtension($floorOffsetY, $width, $depth, $horizontalAxisSlope)
    {
        $widthAxisY = $width * $horizontalAxisSlope * 0.5;
        $depthAxisY = $depth->Y * 0.5;
        $verticalRadius = (float)sqrt(
            $widthAxisY * $widthAxisY + $depthAxisY * $depthAxisY);
        return max(0.0, $floorOffsetY + $depthAxisY - $verticalRadius);
    }

    public static function Draw($canvas, $bottomCenter, $width, $heightOffset, $depth, $horizontalAxisSlope, $series, $drawTopFace)
    {
        if ($canvas === null || $series === null || $width <= 0.0) {
            return;
        }

        $widthAxis = new SKPoint($width * 0.5, $width * $horizontalAxisSlope * 0.5);
        $depthAxis = new SKPoint($depth->X * 0.5, $depth->Y * 0.5);
        $bottom = self::Offset($bottomCenter, $depthAxis);
        $top = self::Offset($bottom, $heightOffset);
        $extremeAngle = self::ResolveHorizontalExtremeAngle($widthAxis, $depthAxis);
        $left = self::EllipsePoint($top, $widthAxis, $depthAxis, $extremeAngle + M_PI);
        $right = self::EllipsePoint($top, $widthAxis, $depthAxis, $extremeAngle);

        $colors = array(
            self::Shade($series->Color, 0.72),
            $series->Color,
            self::Shade($series->Color, 0.72),
        );
        $positions = array(0.0, 0.48, 1.0);
        $shader = SKShader::CreateLinearGradient(
            $left,
            $right,
            $colors,
            $positions,
            SKShaderTileMode::Clamp);
        $bodyPaint = new SKPaint();
        $bodyPaint->Style = SKPaintStyle::Fill;
        $bodyPaint->Shader = $shader;
        $bodyPaint->IsAntialias = true;
        $capPaint = new SKPaint();
        $capPaint->Style = SKPaintStyle::Fill;
        $capPaint->Color = self::Shade($series->Color, 1.07);
        $capPaint->IsAntialias = true;
        $edgePaint = new SKPaint();
        $edgePaint->Style = SKPaintStyle::Stroke;
        $edgePaint->Color = $series->LineColor;
        $edgePaint->StrokeWidth = (float)max(0.25, $series->LineWidthPt);
        $edgePaint->IsAntialias = true;

        $body = self::BuildBodyPath($bottom, $top, $widthAxis, $depthAxis);
        $canvas->DrawPath($body, $bodyPaint);
        if ($series->HasVisibleLine) {
            $canvas->DrawPath($body, $edgePaint);
        }
        $body->Dispose();

        if ($drawTopFace) {
            $cap = self::BuildEllipsePath($top, $widthAxis, $depthAxis);
            $canvas->DrawPath($cap, $capPaint);
            if ($series->HasVisibleLine) {
                $canvas->DrawPath($cap, $edgePaint);
            }
            $cap->Dispose();
        }

        $edgePaint->Dispose();
        $capPaint->Dispose();
        $bodyPaint->Dispose();
        $shader->Dispose();
    }

    private static function BuildBodyPath($bottom, $top, $widthAxis, $depthAxis)
    {
        $path = new SKPath();
        $startAngle = self::ResolveHorizontalExtremeAngle($widthAxis, $depthAxis) + M_PI;
        for ($index = 0; $index <= self::ArcSteps; $index++) {
            $angle = $startAngle + M_PI * $index / self::ArcSteps;
            $point = self::EllipsePoint($top, $widthAxis, $depthAxis, $angle);
            if ($index === 0) {
                $path->MoveTo($point);
            } else {
                $path->LineTo($point);
            }
        }

        for ($index = self::ArcSteps; $index >= 0; $index--) {
            $angle = $startAngle + M_PI * $index / self::ArcSteps;
            $path->LineTo(self::EllipsePoint($bottom, $widthAxis, $depthAxis, $angle));
        }
        $path->Close();
        return $path;
    }

    private static function ResolveHorizontalExtremeAngle($widthAxis, $depthAxis)
    {
        return atan2($depthAxis->X, $widthAxis->X);
    }

    private static function BuildEllipsePath($center, $widthAxis, $depthAxis)
    {
        $path = new SKPath();
        $steps = self::ArcSteps * 2;
        for ($index = 0; $index <= $steps; $index++) {
            $angle = M_PI * 2.0 * $index / $steps;
            $point = self::EllipsePoint($center, $widthAxis, $depthAxis, $angle);
            if ($index === 0) {
                $path->MoveTo($point);
            } else {
                $path->LineTo($point);
            }
        }
        $path->Close();
        return $path;
    }

    private static function EllipsePoint($center, $widthAxis, $depthAxis, $angle)
    {
        $cosine = (float)cos($angle);
        $sine = (float)sin($angle);
        return new SKPoint(
            $center->X + $widthAxis->X * $cosine + $depthAxis->X * $sine,
            $center->Y + $widthAxis->Y * $cosine + $depthAxis->Y * $sine);
    }

    private static function Offset($point, $offset)
    {
        return new SKPoint($point->X + $offset->X, $point->Y + $offset->Y);
    }

    private static function Shade($color, $factor)
    {
        $red = max(0.0, min(255.0, $color->Red * $factor));
        $green = max(0.0, min(255.0, $color->Green * $factor));
        $blue = max(0.0, min(255.0, $color->Blue * $factor));
        return new SKColor((int)round($red), (int)round($green),
            (int)round($blue), $color->Alpha);
    }
}
