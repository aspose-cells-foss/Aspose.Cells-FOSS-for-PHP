<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;

final class MapChartRenderer
{
    private $context;

    public function __construct($context)
    {
        $this->context = $context;
    }

    public function Draw($canvas, $area, $chart, $page = null)
    {
        if ($chart === null || $area->Width <= 0.0 || $area->Height <= 0.0) {
            return;
        }

        $canvas->Save();
        $canvas->ClipRect($area);
        $border = new SKPaint();
        $border->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $border->Style = SKPaintStyle::Stroke;
        $border->StrokeWidth = 0.55;
        $canvas->DrawRect($area, $border);
        $border->Dispose();

        $this->DrawCenteredText($canvas, $page, $area, $chart->Title, $area->MidX, $area->Top + $area->Height * 0.078,
            14.0, new SKColor(0x59, 0x59, 0x59));

        $legendFont = $this->CreateTextPaint(8.0, new SKColor(0x59, 0x59, 0x59));
        $legendNameWidth = $legendFont->MeasureText($chart->SeriesName !== null ? $chart->SeriesName : '');
        $legendFont->Dispose();

        $compactLegend = max(0.0, min(1.0, (70.0 - $legendNameWidth) / 53.0));
        $map = new SKRect($area->Left + $area->Width * (0.11 + 0.045 * $compactLegend),
            $area->Top + $area->Height * (0.296 - 0.066 * $compactLegend),
            $area->Left + $area->Width * (0.678 + 0.197 * $compactLegend),
            $area->Top + $area->Height * (0.81 + 0.055 * $compactLegend));
        $this->DrawCountries($canvas, $map, $chart);
        $this->DrawLegend($canvas, $page, $area, $chart, $compactLegend);
        $this->DrawCenteredText($canvas, $page, $area, 'Map data: Natural Earth', $area->MidX,
            $area->Bottom - $area->Height * 0.043, 5.5, new SKColor(0x66, 0x66, 0x66));
        $canvas->Restore();
    }

    private function DrawCountries($canvas, $map, $chart)
    {
        $stream = null;
        $resourcePath = __DIR__ . '/MapWorld110m.txt';
        if (is_file($resourcePath)) {
            $stream = fopen($resourcePath, 'r');
        }
        if ($stream === false || $stream === null) {
            return;
        }

        $fill = new SKPaint();
        $fill->IsAntialias = true;
        $fill->Style = SKPaintStyle::Fill;
        $outline = new SKPaint();
        $outline->IsAntialias = true;
        $outline->Style = SKPaintStyle::Stroke;
        $outline->StrokeWidth = 0.18;
        $outline->Color = SKColors::White();
        while (($line = fgets($stream)) !== false) {
            $line = rtrim($line, "\r\n");
            if (strlen($line) === 0 || $line[0] === '#') {
                continue;
            }

            $separator = strpos($line, '|');
            if ($separator === false || $separator < 0) {
                continue;
            }

            $value = 0.0;
            $country = substr($line, 0, $separator);
            $matched = self::TryGetValue($chart->Values, $country, $value);
            if (!$matched && strcasecmp($country, 'United States of America') === 0) {
                $matched = self::TryGetValue($chart->Values, 'United States', $value);
            }

            $fill->Color = $matched ? self::ValueColor($value, $chart)
                : new SKColor(0xDF, 0xDF, 0xDF);
            $path = self::BuildPath(substr($line, $separator + 1), $map);
            if ($path === null) {
                continue;
            }

            $canvas->DrawPath($path, $fill);
            $canvas->DrawPath($path, $outline);
            $path->Dispose();
        }

        $outline->Dispose();
        $fill->Dispose();
        fclose($stream);
    }

    private static function TryGetValue($values, $key, &$value)
    {
        if (array_key_exists($key, $values)) {
            $value = $values[$key];
            return true;
        }

        foreach ($values as $candidateKey => $candidateValue) {
            if (strcasecmp((string)$candidateKey, $key) === 0) {
                $value = $candidateValue;
                return true;
            }
        }

        return false;
    }

    private static function BuildPath($coordinates, $map)
    {
        $points = explode(' ', $coordinates);
        if (count($points) < 3) {
            return null;
        }

        $path = new SKPath();
        $count = 0;
        for ($i = 0; $i < count($points); $i++) {
            $comma = strpos($points[$i], ',');
            if ($comma <= 0) {
                continue;
            }

            $longitude = 0.0;
            $latitude = 0.0;
            $longitudeText = substr($points[$i], 0, $comma);
            $latitudeText = substr($points[$i], $comma + 1);
            if (!is_numeric($longitudeText) || !is_numeric($latitudeText)) {
                continue;
            }
            $longitude = (float)$longitudeText;
            $latitude = (float)$latitudeText;

            $x = $map->Left + ($longitude + 180.0) / 360.0 * $map->Width;
            $y = $map->Top + (90.0 - $latitude) / 180.0 * $map->Height;
            if ($count === 0) {
                $path->MoveTo($x, $y);
            } else {
                $path->LineTo($x, $y);
            }

            $count++;
        }

        if ($count < 3) {
            $path->Dispose();
            return null;
        }

        $path->Close();
        return $path;
    }

    private static function ValueColor($value, $chart)
    {
        $fraction = $chart->Maximum > $chart->Minimum
            ? ($value - $chart->Minimum) / ($chart->Maximum - $chart->Minimum) : 0.5;
        $fraction = max(0.0, min(1.0, $fraction));
        if ($chart->HasCustomColors) {
            if ($chart->HasMiddleColor) {
                return $fraction <= 0.5
                    ? self::Blend($chart->MinimumColor, $chart->MiddleColor, $fraction * 2.0)
                    : self::Blend($chart->MiddleColor, $chart->MaximumColor, ($fraction - 0.5) * 2.0);
            }

            return self::Blend($chart->MinimumColor, $chart->MaximumColor, $fraction);
        }

        return new SKColor((int)(209.0 - 175.0 * $fraction),
            (int)(218.0 - 137.0 * $fraction),
            (int)(237.0 - 80.0 * $fraction));
    }

    private static function Blend($first, $second, $fraction)
    {
        return new SKColor(
            (int)round($first->Red + ($second->Red - $first->Red) * $fraction),
            (int)round($first->Green + ($second->Green - $first->Green) * $fraction),
            (int)round($first->Blue + ($second->Blue - $first->Blue) * $fraction));
    }

    private function DrawLegend($canvas, $page, $area, $chart, $compactLegend)
    {
        $x = $area->Left + $area->Width * (0.737 + 0.173 * $compactLegend);
        $top = $area->Top + $area->Height * 0.211;
        $this->DrawLeftText($canvas, $page, $area, $chart->SeriesName, $x - 3.0, $top, 8.0);

        $bar = new SKRect($x, $top + $area->Height * 0.047,
            $x + $area->Width * 0.022, $top + $area->Height * 0.32);
        $paint = new SKPaint();
        $paint->Shader = SKShader::CreateLinearGradient(new SKPoint($bar->MidX, $bar->Top),
            new SKPoint($bar->MidX, $bar->Bottom),
            $chart->HasMiddleColor
                ? array(self::ValueColor($chart->Maximum, $chart), self::ValueColor(($chart->Minimum + $chart->Maximum) * 0.5, $chart), self::ValueColor($chart->Minimum, $chart))
                : array(self::ValueColor($chart->Maximum, $chart), self::ValueColor($chart->Minimum, $chart)),
            $chart->HasMiddleColor ? array(0.0, 0.5, 1.0) : null,
            SKShaderTileMode::Clamp);
        $canvas->DrawRect($bar, $paint);
        $paint->Dispose();

        $this->DrawLeftText($canvas, $page, $area, $chart->MaximumLabel === null || $chart->MaximumLabel === ''
            ? self::FormatPercent($chart->Maximum) : $chart->MaximumLabel,
            $bar->Right + 3.0, $bar->Top + 3.0, 8.0);
        $this->DrawLeftText($canvas, $page, $area, $chart->MinimumLabel === null || $chart->MinimumLabel === ''
            ? self::FormatPercent($chart->Minimum) : $chart->MinimumLabel,
            $bar->Right + 3.0, $bar->Bottom + 3.0, 8.0);
    }

    private static function FormatPercent($value)
    {
        return number_format($value * 100, 3, '.', '') . '%';
    }

    private function DrawCenteredText($canvas, $page, $clip, $text, $centerX, $baselineY,
        $size, $color)
    {
        if ($text === null || $text === '') {
            return;
        }

        $paint = $this->CreateTextPaint($size, $color);
        $this->DrawText($canvas, $page, $clip, $text, $centerX - $paint->MeasureText($text) * 0.5, $baselineY, $paint);
        $paint->Dispose();
    }

    private function DrawLeftText($canvas, $page, $clip, $text, $x, $baselineY, $size)
    {
        $paint = $this->CreateTextPaint($size, new SKColor(0x59, 0x59, 0x59));
        $this->DrawText($canvas, $page, $clip, $text, $x, $baselineY, $paint);
        $paint->Dispose();
    }

    private function DrawText($canvas, $page, $clip, $text,
        $x, $baselineY, $paint)
    {
        $session = $this->context->PdfTextSession;
        if ($session !== null && $page !== null) {
            $session->RecordChartRun($page, $clip, $text, $paint->Typeface, $paint->Color,
                $paint->TextSize, $x, $baselineY);
            return;
        }

        $canvas->DrawText($text, $x, $baselineY, $paint);
    }

    private function CreateTextPaint($size, $color)
    {
        $font = new FontValue();
        $font->Name = 'Calibri';
        $font->Size = $size;
        $paint = new SKPaint();
        $paint->IsAntialias = true;
        $paint->Typeface = $this->context->Fonts->Resolve($font);
        $paint->TextSize = $size;
        $paint->Color = $color;
        $paint->Style = SKPaintStyle::Fill;
        return $paint;
    }
}
