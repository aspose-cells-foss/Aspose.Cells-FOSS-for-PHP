<?php

namespace Aspose\Cells_FOSS\Rendering;

/**
 * Number formatting helpers shared by the PDF post-processing pipeline.
 */
final class PdfNumbers
{
    private function __construct()
    {
    }

    /** "0.###" invariant-culture formatting used by font/text builders. */
    public static function threeDecimals($value)
    {
        $value = (float)$value;
        if ($value == 0.0) {
            return '0';
        }

        $rounded = round($value, 3);
        if ($rounded == 0.0) {
            return '0';
        }

        return self::trimTrailingZeros(number_format($rounded, 3, '.', ''));
    }

    /** "0.#" with 1-decimal rounding used by PdfPathSerializer. */
    public static function oneDecimal($value)
    {
        $value = (float)$value;
        $rounded = round($value, 1);
        if ($rounded == 0.0) {
            return '0';
        }

        return self::trimTrailingZeros(number_format($rounded, 1, '.', ''));
    }

    /** Compact decimal used by PdfContentStreamOptimizer: 3 decimals, no leading zero. */
    public static function compact($value)
    {
        $value = (float)$value;
        $rounded = round($value, 3);
        if ($rounded == 0.0) {
            return '0';
        }

        $text = self::trimTrailingZeros(number_format($rounded, 3, '.', ''));
        if (strncmp($text, '0.', 2) === 0) {
            return substr($text, 1);
        }

        if (strncmp($text, '-0.', 3) === 0) {
            return '-' . substr($text, 2);
        }

        return $text;
    }

    private static function trimTrailingZeros($text)
    {
        if (strpos($text, '.') !== false) {
            $text = rtrim($text, '0');
            $text = rtrim($text, '.');
        }

        if ($text === '' || $text === '-') {
            return '0';
        }

        return $text;
    }
}
