<?php

namespace Aspose\Cells_FOSS;

final class StyleValueSanitizer
{
    private function __construct()
    {
    }

    public static function NormalizeIndentLevel($value)
    {
        if ($value === null || $value < 0 || $value > 250) {
            return 0;
        }

        return $value;
    }

    public static function NormalizeTextRotation($value)
    {
        if ($value === null) {
            return 0;
        }

        if ($value === 255) {
            return 255;
        }

        if ($value < 0 || $value > 180) {
            return 0;
        }

        return $value;
    }

    public static function NormalizeReadingOrder($value)
    {
        if ($value === null || $value < 0 || $value > 2) {
            return 0;
        }

        return $value;
    }
}
