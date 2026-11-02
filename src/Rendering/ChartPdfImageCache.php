<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKBitmap;
use Aspose\Cells_FOSS\Skia\SKCanvas;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKData;
use Aspose\Cells_FOSS\Skia\SKEncodedImageFormat;
use Aspose\Cells_FOSS\Skia\SKImage;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKRectI;

final class ChartPdfImageCache
{
    private $dataByKey = array();

    private $imagesByKey = array();

    public function GetImage($chart, $parsed, $rect, $visualRect, $renderer)
    {
        if ($chart === null || $parsed === null || $renderer === null || $rect->Width <= 0.0 || $rect->Height <= 0.0) {
            return null;
        }

        $targetDpi = self::ResolveTargetDpi($parsed, $rect);
        $encodeLosslessly = self::ShouldEncodeLosslessly($parsed);
        $jpegQuality = self::ResolveJpegQuality($parsed, $rect, $targetDpi);
        $widthPx = self::ToTargetPixels($rect->Width, $targetDpi);
        $heightPx = self::ToTargetPixels($rect->Height, $targetDpi);
        $cropBounds = self::ResolveCropBounds($rect, $visualRect, $widthPx, $heightPx);
        $key = self::BuildKey($chart, $widthPx, $heightPx, $jpegQuality, $encodeLosslessly, $cropBounds);

        if (array_key_exists($key, $this->imagesByKey)) {
            return $this->imagesByKey[$key];
        }

        $bitmap = new SKBitmap($widthPx, $heightPx, true);
        $canvas = new SKCanvas($bitmap);
        try {
            $canvas->Clear(SKColors::White());
            $scaleX = $widthPx / $rect->Width;
            $scaleY = $heightPx / $rect->Height;
            $canvas->Scale($scaleX, $scaleY);
            $renderer->Draw($canvas, new SKRect(0.0, 0.0, $rect->Width, $rect->Height), $parsed);

            $image = SKImage::FromBitmap($bitmap);
            if ($image === null) {
                return null;
            }

            try {
                $croppedBitmap = self::CropBitmap($bitmap, $cropBounds);
                try {
                    $croppedImage = SKImage::FromBitmap($croppedBitmap);
                    if ($croppedImage === null) {
                        return null;
                    }

                    if ($encodeLosslessly) {
                        $encoded = $croppedImage->Encode(SKEncodedImageFormat::Png, 100);
                    } else {
                        $encoded = $croppedImage->Encode(SKEncodedImageFormat::Jpeg, $jpegQuality);
                    }

                    $croppedImage->Dispose();
                } finally {
                    $croppedBitmap->Dispose();
                }

                if ($encoded === null || $encoded->Size <= 0) {
                    return null;
                }

                $cached = SKImage::FromEncodedData($encoded);
                if ($cached === null) {
                    $encoded->Dispose();
                    return null;
                }

                $this->dataByKey[$key] = $encoded;
                $this->imagesByKey[$key] = $cached;
                return $cached;
            } finally {
                $image->Dispose();
            }
        } finally {
            $canvas->Dispose();
            $bitmap->Dispose();
        }
    }

    private static function ToTargetPixels($points, $targetDpi)
    {
        $pixels = $points * $targetDpi / 72.0;
        if ($pixels < 1.0) {
            $pixels = 1.0;
        }

        return (int)ceil($pixels);
    }

    private static function ResolveTargetDpi($parsed, $rect)
    {
        $area = $rect->Width * $rect->Height;
        if ($parsed !== null && ($parsed->Kind === ChartKind::Line || $parsed->Kind === ChartKind::Area)) {
            $pointCount = self::MaxPointCount($parsed);
            if ($pointCount <= 12 && $area >= 50000.0) {
                return 192.0;
            }

            return 216.0;
        }

        if ($parsed !== null && ($parsed->Kind === ChartKind::Column || $parsed->Kind === ChartKind::Bar)) {
            if ($area >= 50000.0) {
                return 180.0;
            }

            return 192.0;
        }

        if ($area >= 70000.0) {
            return 144.0;
        }

        return 168.0;
    }

    private static function ResolveJpegQuality($parsed, $rect, $targetDpi)
    {
        if ($parsed !== null && $parsed->Kind === ChartKind::Line) {
            if ($targetDpi <= 120.0) {
                return 58;
            }

            return 66;
        }

        if ($parsed !== null && $parsed->Kind === ChartKind::Area) {
            return $targetDpi <= 150.0 ? 68 : 74;
        }

        return $targetDpi <= 144.0 ? 70 : 76;
    }

    private static function ShouldEncodeLosslessly($parsed)
    {
        if ($parsed === null) {
            return false;
        }

        return $parsed->Kind === ChartKind::Line
            || $parsed->Kind === ChartKind::Area
            || $parsed->Kind === ChartKind::Column
            || $parsed->Kind === ChartKind::Bar;
    }

    private static function MaxPointCount($chart)
    {
        $count = $chart !== null ? count($chart->Categories) : 0;
        if ($chart === null) {
            return $count;
        }

        for ($i = 0; $i < count($chart->Series); $i++) {
            if ($chart->Series[$i] !== null && count($chart->Series[$i]->Values) > $count) {
                $count = count($chart->Series[$i]->Values);
            }
        }

        return $count;
    }

    private static function ResolveCropBounds($rect, $visualRect, $widthPx, $heightPx)
    {
        if ($visualRect->Width <= 0.0 || $visualRect->Height <= 0.0) {
            return new SKRectI(0, 0, $widthPx, $heightPx);
        }

        $scaleX = $widthPx / $rect->Width;
        $scaleY = $heightPx / $rect->Height;
        $left = (int)floor(($visualRect->Left - $rect->Left) * $scaleX);
        $top = (int)floor(($visualRect->Top - $rect->Top) * $scaleY);
        $right = (int)ceil(($visualRect->Right - $rect->Left) * $scaleX);
        $bottom = (int)ceil(($visualRect->Bottom - $rect->Top) * $scaleY);

        if ($left < 0) {
            $left = 0;
        }

        if ($top < 0) {
            $top = 0;
        }

        if ($right > $widthPx) {
            $right = $widthPx;
        }

        if ($bottom > $heightPx) {
            $bottom = $heightPx;
        }

        if ($right <= $left) {
            $left = 0;
            $right = $widthPx;
        }

        if ($bottom <= $top) {
            $top = 0;
            $bottom = $heightPx;
        }

        return new SKRectI($left, $top, $right, $bottom);
    }

    private static function CropBitmap($bitmap, $cropBounds)
    {
        $width = $cropBounds->Right - $cropBounds->Left;
        $height = $cropBounds->Bottom - $cropBounds->Top;
        if ($width <= 0 || $height <= 0) {
            return $bitmap->Copy();
        }

        $result = new SKBitmap($width, $height, true);
        $canvas = new SKCanvas($result);
        try {
            $canvas->Clear(SKColors::White());
            $source = new SKRect($cropBounds->Left, $cropBounds->Top, $cropBounds->Right, $cropBounds->Bottom);
            $destination = new SKRect(0.0, 0.0, $width, $height);
            $canvas->DrawBitmap($bitmap, $source, $destination);
        } finally {
            $canvas->Dispose();
        }

        return $result;
    }

    private static function BuildKey($chart, $widthPx, $heightPx, $jpegQuality, $encodeLosslessly, $cropBounds)
    {
        $hash = 17;
        $hash = $hash * 31 + ($chart->RawChartXml !== null ? crc32($chart->RawChartXml) : 0);
        $hash = $hash * 31 + $chart->UpperLeftRow;
        $hash = $hash * 31 + $chart->UpperLeftColumn;
        $hash = $hash * 31 + $chart->LowerRightRow;
        $hash = $hash * 31 + $chart->LowerRightColumn;
        $hash = $hash * 31 + $widthPx;
        $hash = $hash * 31 + $heightPx;
        $hash = $hash * 31 + $jpegQuality;
        $hash = $hash * 31 + ($encodeLosslessly ? 1 : 0);
        $hash = $hash * 31 + $cropBounds->Left;
        $hash = $hash * 31 + $cropBounds->Top;
        $hash = $hash * 31 + $cropBounds->Right;
        $hash = $hash * 31 + $cropBounds->Bottom;
        return (string)$hash;
    }

    public function Dispose()
    {
        foreach ($this->imagesByKey as $image) {
            if ($image !== null) {
                $image->Dispose();
            }
        }

        foreach ($this->dataByKey as $data) {
            if ($data !== null) {
                $data->Dispose();
            }
        }

        $this->imagesByKey = array();
        $this->dataByKey = array();
    }
}
