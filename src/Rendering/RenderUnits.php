<?php

namespace Aspose\Cells_FOSS\Rendering;

final class RenderUnits
{
    const PointsPerInch = 72.0;

    const PixelsToPoints = 72.0 / 96.0;

    const PointsToPixels = 96.0 / 72.0;

    private function __construct()
    {
    }

    public static function InchesToPoints($inches)
    {
        return $inches * self::PointsPerInch;
    }

    public static function PixelsToPointsValue($pixels)
    {
        return $pixels * self::PixelsToPoints;
    }

    public static function ColumnWidthCharsToPixels($widthChars, $maxDigitWidth)
    {
        if ($maxDigitWidth <= 0.0) {
            $maxDigitWidth = 7.5;
        }

        return (256.0 * $widthChars + intval(128.0 / $maxDigitWidth)) / 256.0 * $maxDigitWidth;
    }

    public static function ColumnWidthCharsToPoints($widthChars, $maxDigitWidth)
    {
        return self::PixelsToPointsValue(self::ColumnWidthCharsToPixels($widthChars, $maxDigitWidth));
    }

    public static function DefaultColumnWidthCharsToPoints($displayChars, $maxDigitWidth)
    {
        if ($maxDigitWidth <= 0.0) {
            $maxDigitWidth = 7.5;
        }

        $pixels = round($displayChars * $maxDigitWidth, 0, PHP_ROUND_HALF_UP) + 4.0;
        return self::PixelsToPointsValue($pixels);
    }
}
