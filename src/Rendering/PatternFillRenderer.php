<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Skia\SKCanvas;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKRect;

/**
 * Draws classic SpreadsheetML pattern fills using foreground/background colors and a small
 * tiled geometry, instead of collapsing every non-solid fill to a solid foreground color.
 */
final class PatternFillRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $rect, $style, $colors)
    {
        if ($canvas === null || $colors === null || $style === null || $rect->Width <= 0.0 || $rect->Height <= 0.0) {
            return false;
        }

        if ($style->Pattern === FillPatternKind::None) {
            return false;
        }

        if ($style->Pattern === FillPatternKind::Solid) {
            $solid = $colors->Resolve($style->ForegroundColor, SKColors::Transparent());
            if ($solid->Alpha === 0) {
                $solid = $colors->Resolve($style->BackgroundColor, SKColors::Transparent());
            }

            if ($solid->Alpha === 0) {
                return false;
            }

            $paint = new SKPaint();
            try {
                $paint->Style = SKPaintStyle::Fill;
                $paint->IsAntialias = false;
                $paint->Color = $solid;
                $canvas->DrawRect($rect, $paint);
            } finally {
                $paint->Dispose();
            }

            return true;
        }

        $background = $colors->Resolve($style->BackgroundColor, SKColors::White());
        $foreground = $colors->Resolve($style->ForegroundColor, SKColors::Black());

        $backgroundPaint = new SKPaint();
        try {
            $backgroundPaint->Style = SKPaintStyle::Fill;
            $backgroundPaint->IsAntialias = false;
            $backgroundPaint->Color = $background;
            $canvas->DrawRect($rect, $backgroundPaint);
        } finally {
            $backgroundPaint->Dispose();
        }

        if ($foreground->Alpha === 0) {
            return $background->Alpha !== 0;
        }

        $canvas->Save();
        $canvas->ClipRect($rect);

        switch ($style->Pattern) {
            case FillPatternKind::MediumGray:
                self::DrawDensityPattern($canvas, $rect, $foreground, $background, 0.40, 2.5, 2.5, 0.15, 50, true);
                break;
            case FillPatternKind::DarkGray:
                self::DrawDensityHolePattern($canvas, $rect, $foreground, $background, 0.64, 2.35, 2.35, 0.10, 60, true);
                break;
            case FillPatternKind::LightGray:
                self::DrawDensityPattern($canvas, $rect, $foreground, $background, 0.145, 2.9, 2.9, 0.10, 36, true);
                break;
            case FillPatternKind::Gray125:
                self::DrawDensityPattern($canvas, $rect, $foreground, $background, 0.065, 3.45, 3.45, 0.10, 38, true);
                break;
            case FillPatternKind::Gray0625:
                self::DrawDensityPattern($canvas, $rect, $foreground, $background, 0.028, 4.6, 4.6, 0.09, 36, true);
                break;
            case FillPatternKind::DarkHorizontal:
                self::DrawStripedPattern($canvas, $rect, $foreground, $background, 0.00, 1.0, 0.34, true, 255);
                break;
            case FillPatternKind::DarkVertical:
                self::DrawStripedPattern($canvas, $rect, $foreground, $background, 0.16, 1.56, 0.66, false, 244);
                break;
            case FillPatternKind::DarkDown:
                self::DrawDiagonalPattern($canvas, $rect, $foreground, $background, 2.3, 0.34, true, 154);
                break;
            case FillPatternKind::DarkUp:
                self::DrawDiagonalPattern($canvas, $rect, $foreground, $background, 2.3, 0.34, false, 154);
                break;
            case FillPatternKind::DarkGrid:
                self::DrawGridPattern($canvas, $rect, $foreground, $background, 2.3, 0.32, 146);
                break;
            case FillPatternKind::DarkTrellis:
                self::DrawTrellisPattern($canvas, $rect, $foreground, $background, 2.5, 0.32, 146);
                break;
            case FillPatternKind::LightHorizontal:
                self::DrawStripedPattern($canvas, $rect, $foreground, $background, 0.045, 3.0, 0.24, true, 136);
                break;
            case FillPatternKind::LightVertical:
                self::DrawStripedPattern($canvas, $rect, $foreground, $background, 0.045, 3.0, 0.24, false, 136);
                break;
            case FillPatternKind::LightDown:
                self::DrawDiagonalPattern($canvas, $rect, $foreground, $background, 3.3, 0.17, true, 102);
                break;
            case FillPatternKind::LightUp:
                self::DrawDiagonalPattern($canvas, $rect, $foreground, $background, 3.3, 0.17, false, 102);
                break;
            case FillPatternKind::LightGrid:
                self::DrawGridPattern($canvas, $rect, $foreground, $background, 3.2, 0.17, 96);
                break;
            case FillPatternKind::LightTrellis:
                self::DrawTrellisPattern($canvas, $rect, $foreground, $background, 1.05, 0.16, 136);
                break;
        }

        $canvas->Restore();
        return true;
    }

    private static function DrawDensityPattern($canvas, $rect, $foreground, $background, $coverage, $stepX, $stepY, $dotSize, $alpha, $staggered)
    {
        self::FillRect($canvas, $rect, self::Blend($background, $foreground, $coverage));
        self::DrawDots($canvas, $rect, self::WithAlpha($foreground, $alpha), $stepX, $stepY, $dotSize, $staggered);
    }

    private static function DrawDensityHolePattern($canvas, $rect, $foreground, $background, $coverage, $stepX, $stepY, $holeSize, $alpha, $staggered)
    {
        self::FillRect($canvas, $rect, self::Blend($background, $foreground, $coverage));
        self::DrawDots($canvas, $rect, self::WithAlpha($background, $alpha), $stepX, $stepY, $holeSize, $staggered);
    }

    private static function DrawStripedPattern($canvas, $rect, $foreground, $background, $coverage, $step, $strokeWidth, $horizontal, $alpha)
    {
        self::FillRect($canvas, $rect, self::Blend($background, $foreground, $coverage));
        $accent = self::WithAlpha($foreground, $alpha);
        if ($horizontal) {
            self::DrawHorizontalLines($canvas, $rect, $accent, $step, $strokeWidth);
        } else {
            self::DrawVerticalLines($canvas, $rect, $accent, $step, $strokeWidth);
        }
    }

    private static function DrawDiagonalPattern($canvas, $rect, $foreground, $background, $step, $strokeWidth, $down, $alpha)
    {
        self::FillRect($canvas, $rect, $background);
        self::DrawDiagonalLines($canvas, $rect, self::WithAlpha($foreground, $alpha), $step, $strokeWidth, $down);
    }

    private static function DrawGridPattern($canvas, $rect, $foreground, $background, $step, $strokeWidth, $alpha)
    {
        self::FillRect($canvas, $rect, $background);
        $accent = self::WithAlpha($foreground, $alpha);
        self::DrawHorizontalLines($canvas, $rect, $accent, $step, $strokeWidth);
        self::DrawVerticalLines($canvas, $rect, $accent, $step, $strokeWidth);
    }

    private static function DrawTrellisPattern($canvas, $rect, $foreground, $background, $step, $strokeWidth, $alpha)
    {
        self::FillRect($canvas, $rect, $background);
        $accent = self::WithAlpha($foreground, $alpha);
        self::DrawDiagonalLines($canvas, $rect, $accent, $step, $strokeWidth, true);
        self::DrawDiagonalLines($canvas, $rect, $accent, $step, $strokeWidth, false);
    }

    private static function FillRect($canvas, $rect, $color)
    {
        $paint = new SKPaint();
        try {
            $paint->Style = SKPaintStyle::Fill;
            $paint->IsAntialias = false;
            $paint->Color = $color;
            $canvas->DrawRect($rect, $paint);
        } finally {
            $paint->Dispose();
        }
    }

    private static function DrawDots($canvas, $rect, $color, $stepX, $stepY, $dotSize, $staggered)
    {
        $paint = new SKPaint();
        try {
            $paint->Style = SKPaintStyle::Fill;
            $paint->IsAntialias = false;
            $paint->Color = $color;

            $rowIndex = 0;
            for ($top = $rect->Top; $top < $rect->Bottom; $top += $stepY) {
                $offset = $staggered && ($rowIndex % 2) === 1 ? $stepX * 0.5 : 0.0;
                for ($left = $rect->Left + $offset; $left < $rect->Right; $left += $stepX) {
                    $canvas->DrawRect($left, $top, $dotSize, $dotSize, $paint);
                }

                $rowIndex++;
            }
        } finally {
            $paint->Dispose();
        }
    }

    private static function DrawHorizontalLines($canvas, $rect, $color, $step, $strokeWidth)
    {
        $paint = new SKPaint();
        try {
            $paint->Style = SKPaintStyle::Stroke;
            $paint->IsAntialias = false;
            $paint->Color = $color;
            $paint->StrokeWidth = $strokeWidth;

            for ($y = $rect->Top + $strokeWidth * 0.5; $y < $rect->Bottom; $y += $step) {
                $canvas->DrawLine($rect->Left, $y, $rect->Right, $y, $paint);
            }
        } finally {
            $paint->Dispose();
        }
    }

    private static function DrawVerticalLines($canvas, $rect, $color, $step, $strokeWidth)
    {
        $paint = new SKPaint();
        try {
            $paint->Style = SKPaintStyle::Stroke;
            $paint->IsAntialias = false;
            $paint->Color = $color;
            $paint->StrokeWidth = $strokeWidth;

            for ($x = $rect->Left + $strokeWidth * 0.5; $x < $rect->Right; $x += $step) {
                $canvas->DrawLine($x, $rect->Top, $x, $rect->Bottom, $paint);
            }
        } finally {
            $paint->Dispose();
        }
    }

    private static function DrawDiagonalLines($canvas, $rect, $color, $step, $strokeWidth, $down)
    {
        $paint = new SKPaint();
        try {
            $paint->Style = SKPaintStyle::Stroke;
            $paint->IsAntialias = false;
            $paint->Color = $color;
            $paint->StrokeWidth = $strokeWidth;

            $width = $rect->Width;
            $height = $rect->Height;
            for ($offset = -$height; $offset <= $width; $offset += $step) {
                if ($down) {
                    $canvas->DrawLine($rect->Left + $offset, $rect->Top, $rect->Left + $offset + $height, $rect->Bottom, $paint);
                } else {
                    $canvas->DrawLine($rect->Left + $offset, $rect->Bottom, $rect->Left + $offset + $height, $rect->Top, $paint);
                }
            }
        } finally {
            $paint->Dispose();
        }
    }

    private static function Blend($background, $foreground, $foregroundFraction)
    {
        $clamped = max(0.0, min(1.0, $foregroundFraction));
        $inverse = 1.0 - $clamped;
        return new SKColor(
            (int)round($background->Red * $inverse + $foreground->Red * $clamped, 0, PHP_ROUND_HALF_EVEN),
            (int)round($background->Green * $inverse + $foreground->Green * $clamped, 0, PHP_ROUND_HALF_EVEN),
            (int)round($background->Blue * $inverse + $foreground->Blue * $clamped, 0, PHP_ROUND_HALF_EVEN),
            (int)round($background->Alpha * $inverse + $foreground->Alpha * $clamped, 0, PHP_ROUND_HALF_EVEN));
    }

    private static function WithAlpha($color, $alpha)
    {
        return new SKColor($color->Red, $color->Green, $color->Blue, $alpha);
    }
}
