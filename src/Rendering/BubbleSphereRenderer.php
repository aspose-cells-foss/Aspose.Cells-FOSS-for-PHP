<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKFilterQuality;
use Aspose\Cells_FOSS\Skia\SKImageInfo;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;
use Aspose\Cells_FOSS\Skia\SKSurface;

final class BubbleSphereRenderer
{
    private function __construct()
    {
    }

    public static function CreateTexture($color)
    {
        $surface = SKSurface::Create(new SKImageInfo(96, 96));
        $surface->Canvas->Clear(SKColors::Transparent());
        self::Draw($surface->Canvas, new SKPoint(48.0, 48.0), 46.0, $color);
        $image = $surface->Snapshot();
        $surface->Dispose();
        return $image;
    }

    public static function DrawImage($canvas, $center, $radius, $texture)
    {
        $halfSide = $radius * 48.0 / 46.0;
        $bounds = new SKRect($center->X - $halfSide, $center->Y - $halfSide,
            $center->X + $halfSide, $center->Y + $halfSide);
        $paint = new SKPaint();
        $paint->FilterQuality = SKFilterQuality::High;
        $paint->IsAntialias = true;
        $canvas->DrawImage($texture, $bounds, $paint);
        $paint->Dispose();
    }

    public static function Draw($canvas, $center, $radius, $color)
    {
        if ($radius <= 0.0) {
            return;
        }

        $color->ToHsv($hue, $saturation, $value);
        $highlightHue = $hue < 50.0 ? $hue + 15.0 : $hue;
        $highlightSaturation = $hue < 50.0 ? $saturation * 0.58 : $saturation * 0.68;
        $highlight = SKColor::FromHsv($highlightHue, $highlightSaturation,
            min(100.0, $value + max(9.0, (100.0 - $value) * 0.6)), $color->Alpha);
        $lit = SKColor::FromHsv($hue, $saturation * 0.98,
            min(100.0, $value * 1.08), $color->Alpha);
        $shade = SKColor::FromHsv($hue, $saturation * 0.9, $value * 0.48, $color->Alpha);
        $edge = SKColor::FromHsv($hue, $saturation * 0.85, $value * 0.27, $color->Alpha);
        $lightCenter = new SKPoint($center->X - $radius * 0.20, $center->Y - $radius * 0.44);
        $colors = array($highlight, $lit, $color, $shade, $edge);
        $stops = array(0.0, 0.20, 0.33, 0.65, 1.0);
        $shader = SKShader::CreateRadialGradient($lightCenter, $radius * 1.65,
            $colors, $stops, SKShaderTileMode::Clamp);
        $fill = new SKPaint();
        $fill->Style = SKPaintStyle::Fill;
        $fill->Shader = $shader;
        $fill->IsAntialias = true;
        $reflectedLight = SKShader::CreateLinearGradient(
            $center, new SKPoint($center->X, $center->Y + $radius),
            array(new SKColor(255, 255, 255, 0),
                new SKColor(255, 255, 255, 55)),
            array(0.0, 1.0), SKShaderTileMode::Clamp);
        $reflection = new SKPaint();
        $reflection->Style = SKPaintStyle::Fill;
        $reflection->Shader = $reflectedLight;
        $reflection->IsAntialias = true;
        $specular = SKShader::CreateRadialGradient($lightCenter, $radius * 0.42,
            array(new SKColor(255, 255, 255, 48), SKColors::Transparent()),
            null, SKShaderTileMode::Clamp);
        $glint = new SKPaint();
        $glint->Style = SKPaintStyle::Fill;
        $glint->Shader = $specular;
        $glint->IsAntialias = true;
        $canvas->DrawCircle($center, $radius, $fill);
        $canvas->DrawCircle($center, $radius, $reflection);
        $canvas->DrawCircle($center, $radius, $glint);
        $glint->Dispose();
        $specular->Dispose();
        $reflection->Dispose();
        $reflectedLight->Dispose();
        $fill->Dispose();
        $shader->Dispose();
    }
}
