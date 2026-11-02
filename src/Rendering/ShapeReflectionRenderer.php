<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;

final class ShapeReflectionRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $points, $shape)
    {
        $reflection = $shape->Reflection;
        if ($reflection === null || count($points) < 2) {
            return;
        }
        // Vertical, below-shape reflection. Other reflection transforms are not handled here.
        if (self::Number($reflection, 'sy', -100000.0) >= 0.0 || abs(self::Number($reflection, 'dir', 5400000.0) - 5400000.0) > 1.0) {
            return;
        }
        $top = $points[0]->Y;
        $bottom = $top;
        for ($i = 1; $i < count($points); $i++) {
            $top = min($top, $points[$i]->Y);
            $bottom = max($bottom, $points[$i]->Y);
        }
        $height = $bottom - $top;
        if ($height <= 0.0) {
            return;
        }
        $distance = (float)(self::Number($reflection, 'dist', 0.0) / 12700.0);
        $endPosition = (float)max(0.001, min(0.999, self::Number($reflection, 'endPos', 100000.0) / 100000.0));
        $startColor = $shape->LineColor->WithAlpha(self::Alpha($reflection, 'stA', 100000.0));
        $endColor = $shape->LineColor->WithAlpha(self::Alpha($reflection, 'endA', 0.0));
        $shader = SKShader::CreateLinearGradient(new SKPoint(0.0, $bottom + $distance), new SKPoint(0.0, $bottom + $distance + $height),
            array($startColor, $endColor, $shape->LineColor->WithAlpha(0)), array(0.0, $endPosition, 1.0), SKShaderTileMode::Clamp);
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->StrokeWidth = (float)$shape->LineWidthPt;
        $paint->Shader = $shader;
        $paint->IsAntialias = true;
        $path = new SKPath();
        for ($i = 0; $i < count($points); $i++) {
            $point = new SKPoint($points[$i]->X, 2.0 * $bottom + $distance - $points[$i]->Y);
            if ($i === 0) {
                $path->MoveTo($point);
            } else {
                $path->LineTo($point);
            }
        }
        $canvas->DrawPath($path, $paint);
        $path->Dispose();
        $paint->Dispose();
        $shader->Dispose();
    }

    private static function Alpha($element, $name, $fallback)
    {
        return (int)round(max(0.0, min(100000.0, self::Number($element, $name, $fallback))) * 255.0 / 100000.0, 0, PHP_ROUND_HALF_EVEN);
    }

    private static function Number($element, $name, $fallback)
    {
        $text = (string)$element->Attribute($name);
        if ($text === '' || !is_numeric($text)) {
            return $fallback;
        }

        $value = (float)$text;
        return !is_nan($value) && !is_infinite($value) ? $value : $fallback;
    }
}
