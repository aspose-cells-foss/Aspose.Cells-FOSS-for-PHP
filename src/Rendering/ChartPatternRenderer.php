<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKAlphaType;
use Aspose\Cells_FOSS\Skia\SKBitmap;
use Aspose\Cells_FOSS\Skia\SKClipOperation;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColorType;
use Aspose\Cells_FOSS\Skia\SKMatrix;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;

final class ChartPatternRenderer
{
    private function __construct()
    {
    }

    public static function DrawChartArea($canvas, $rect, $chart)
    {
        if ($canvas === null || $chart === null || $rect->Width <= 0.0 || $rect->Height <= 0.0
            || $chart->ChartAreaPatternPreset !== 'weave') {
            return false;
        }

        $rows = array(
            '...#....##..', '...#....##..', '.##.##....#.', '#.....#....#',
            '#.....#....#', '....##.#..#.', '...#....##..', '...#....##..',
            '.##.##......', '#.....#....#', '#.....#....#', '.##....#..#.'
        );
        $foreground = $chart->ChartAreaPatternForegroundColor;
        $background = $chart->ChartAreaPatternBackgroundColor;
        $tile = new SKBitmap(12, 12, SKColorType::Rgba8888, SKAlphaType::Premul);
        for ($y = 0; $y < count($rows); $y++) {
            for ($x = 0; $x < strlen($rows[$y]); $x++) {
                $source = $rows[$y][$x] === '#' ? $foreground : $background;
                // Near-opaque pixels keep Skia's PDF tile lossless instead of JPEG-compressing away the blue detail.
                $tile->SetPixel($x, $y, new SKColor($source->Red, $source->Green, $source->Blue, 254));
            }
        }

        $shader = $tile->ToShader(SKShaderTileMode::Repeat, SKShaderTileMode::Repeat,
            SKMatrix::CreateScale(0.5, 0.5));
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Fill;
        $paint->Shader = $shader;
        $paint->IsAntialias = false;
        $canvas->DrawRect($rect, $paint);
        $paint->Dispose();
        $shader->Dispose();
        $tile->Dispose();

        return true;
    }

    public static function DrawPath($canvas, $path, $series, $fallbackColor)
    {
        if ($canvas === null || $path === null || $series === null
            || $series->PatternPreset !== 'ltUpDiag') {
            return false;
        }

        $bounds = $path->Bounds;
        $background = $series->PatternBackgroundColor->Alpha > 0
            ? $series->PatternBackgroundColor
            : $fallbackColor;
        $foreground = $series->PatternForegroundColor->Alpha > 0
            ? $series->PatternForegroundColor
            : $series->Color;

        $backgroundPaint = new SKPaint();
        $backgroundPaint->Style = SKPaintStyle::Fill;
        $backgroundPaint->Color = $background;
        $backgroundPaint->IsAntialias = true;
        $canvas->DrawPath($path, $backgroundPaint);
        $backgroundPaint->Dispose();

        $canvas->Save();
        $canvas->ClipPath($path, SKClipOperation::Intersect, true);
        self::DrawLightUpDiagonal($canvas, $bounds, $foreground);
        $canvas->Restore();
        return true;
    }

    public static function DrawRect($canvas, $rect, $series, $fallbackColor)
    {
        if ($canvas === null || $series === null || $rect->Width <= 0.0 || $rect->Height <= 0.0) {
            return false;
        }

        $path = new SKPath();
        $path->AddRect($rect);
        $result = self::DrawPath($canvas, $path, $series, $fallbackColor);
        $path->Dispose();
        return $result;
    }

    private static function DrawLightUpDiagonal($canvas, $bounds, $color)
    {
        $lineColor = new SKColor($color->Red, $color->Green, $color->Blue, 150);
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $lineColor;
        $paint->StrokeWidth = 0.35;
        $paint->IsAntialias = true;
        $height = max(1.0, $bounds->Height);
        for ($x = $bounds->Left - $height; $x <= $bounds->Right; $x += 2.2) {
            $canvas->DrawLine($x, $bounds->Bottom, $x + $height, $bounds->Top, $paint);
        }

        $paint->Dispose();
    }
}
