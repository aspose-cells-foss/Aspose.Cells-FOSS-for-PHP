<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKClipOperation;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKData;
use Aspose\Cells_FOSS\Skia\SKFilterQuality;
use Aspose\Cells_FOSS\Skia\SKImage;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKRect;

final class Area3DWallFillRenderer
{
    private function __construct()
    {
    }

    public static function Draw(
        $canvas,
        $frontBottomLeft,
        $frontTopLeft,
        $backBottomLeft,
        $backTopLeft,
        $backBottomRight,
        $backTopRight,
        $sideWallImageData,
        $backWallImageData,
        $fallbackColor)
    {
        $sideWall = new SKPath();
        $sideWall->MoveTo($frontTopLeft);
        $sideWall->LineTo($backTopLeft);
        $sideWall->LineTo($backBottomLeft);
        $sideWall->LineTo($frontBottomLeft);
        $sideWall->Close();
        self::DrawWall($canvas, $sideWall, $sideWallImageData, $fallbackColor, 0.72);
        $sideWall->Dispose();

        $backWall = new SKPath();
        $backWall->MoveTo($backTopLeft);
        $backWall->LineTo($backTopRight);
        $backWall->LineTo($backBottomRight);
        $backWall->LineTo($backBottomLeft);
        $backWall->Close();
        self::DrawWall($canvas, $backWall, $backWallImageData, $fallbackColor, 1.0);
        $backWall->Dispose();
    }

    public static function DrawFloor(
        $canvas,
        $backLeft,
        $backRight,
        $frontRight,
        $frontLeft,
        $imageData,
        $fallbackColor)
    {
        $floor = new SKPath();
        $floor->MoveTo($backLeft);
        $floor->LineTo($backRight);
        $floor->LineTo($frontRight);
        $floor->LineTo($frontLeft);
        $floor->Close();
        self::DrawWall($canvas, $floor, $imageData, $fallbackColor, 0.88);
        $floor->Dispose();
    }

    private static function DrawWall($canvas, $wall, $imageData, $fallbackColor, $brightness)
    {
        $fallback = new SKPaint();
        $fallback->Style = SKPaintStyle::Fill;
        $fallback->Color = self::Shade($fallbackColor, $brightness);
        $fallback->IsAntialias = true;
        $canvas->DrawPath($wall, $fallback);
        $fallback->Dispose();

        if ($imageData === null || strlen((string)$imageData) === 0) {
            return;
        }

        $data = SKData::CreateCopy($imageData);
        $image = SKImage::FromEncodedData($data);
        self::DrawTiles($canvas, $wall, $image);
        if ($image !== null) {
            $image->Dispose();
        }

        $data->Dispose();

        if ($brightness < 0.999) {
            $alpha = (int)round((1.0 - $brightness) * 255.0);
            $shadow = new SKPaint();
            $shadow->Style = SKPaintStyle::Fill;
            $shadow->Color = new SKColor(0, 0, 0, $alpha);
            $shadow->IsAntialias = true;
            $canvas->DrawPath($wall, $shadow);
            $shadow->Dispose();
        }
    }

    private static function DrawTiles($canvas, $wall, $image)
    {
        if ($image === null || $image->Width <= 0 || $image->Height <= 0) {
            return;
        }

        $bounds = $wall->Bounds;
        $canvas->Save();
        $canvas->ClipPath($wall, SKClipOperation::Intersect, true);
        $paint = new SKPaint();
        $paint->IsAntialias = true;
        $paint->FilterQuality = SKFilterQuality::High;
        $tileWidth = $image->Width * 72.0 / 96.0;
        $tileHeight = $image->Height * 72.0 / 96.0;
        $firstX = $bounds->Left - self::PositiveRemainder($bounds->Left, $tileWidth);
        $firstY = $bounds->Top - self::PositiveRemainder($bounds->Top, $tileHeight);
        for ($y = $firstY; $y < $bounds->Bottom; $y += $tileHeight) {
            for ($x = $firstX; $x < $bounds->Right; $x += $tileWidth) {
                $destination = new SKRect($x, $y, $x + $tileWidth, $y + $tileHeight);
                $canvas->DrawImage($image, $destination, $paint);
            }
        }

        $paint->Dispose();
        $canvas->Restore();
    }

    private static function PositiveRemainder($value, $divisor)
    {
        $remainder = fmod($value, $divisor);
        return $remainder < 0.0 ? $remainder + $divisor : $remainder;
    }

    private static function Shade($color, $brightness)
    {
        return new SKColor(
            (int)round($color->Red * $brightness),
            (int)round($color->Green * $brightness),
            (int)round($color->Blue * $brightness),
            $color->Alpha);
    }
}
