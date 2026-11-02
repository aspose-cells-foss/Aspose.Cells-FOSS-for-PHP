<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\DisplayTextFormatterSupport;
use Aspose\Cells_FOSS\NumberFormat;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class NumberFormatTextColor
{
    private function __construct()
    {
    }

    public static function Resolve($value, $style, $fallback)
    {
        if ($value === null || $style === null) {
            return $fallback;
        }

        $formatCode = NumberFormat::ResolveFormatCode($style->NumberFormat->Number, $style->NumberFormat->Custom);
        if ($formatCode === null || trim($formatCode) === '') {
            return $fallback;
        }

        $sections = DisplayTextFormatterSupport::ParseSections($formatCode);
        $selected = null;
        if (DisplayTextFormatterSupport::IsNumericValue($value)) {
            $useAbsoluteValue = false;
            $selected = DisplayTextFormatterSupport::SelectNumericSection($sections,
                DisplayTextFormatterSupport::ConvertToDouble($value), $useAbsoluteValue);
        } else if ($value instanceof \DateTime) {
            $selected = DisplayTextFormatterSupport::SelectDateTimeSection($sections);
        } else if (is_string($value)) {
            $selected = DisplayTextFormatterSupport::SelectTextSection($sections);
        }

        if ($selected === null) {
            return $fallback;
        }

        $section = $selected->Raw;
        $inQuote = false;
        for ($index = 0; $index < strlen($section); $index++) {
            $character = $section[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($inQuote) {
                continue;
            }

            if ($character === '\\' || $character === '_' || $character === '*') {
                $index++;
                continue;
            }

            if ($character !== '[') {
                continue;
            }

            $end = strpos($section, ']', $index + 1);
            if ($end === false) {
                break;
            }

            $name = substr($section, $index + 1, $end - $index - 1);
            $color = null;
            if (self::TryResolveNamedColor($name, $color)) {
                return $color;
            }

            $index = $end;
        }

        return $fallback;
    }

    private static function TryResolveNamedColor($name, &$color)
    {
        switch (strtolower($name)) {
            case 'black': $color = SKColors::Black(); return true;
            case 'blue': $color = SKColors::Blue(); return true;
            case 'cyan': $color = SKColors::Cyan(); return true;
            case 'green': $color = SKColors::Lime(); return true;
            case 'magenta': $color = SKColors::Magenta(); return true;
            case 'red': $color = SKColors::Red(); return true;
            case 'white': $color = SKColors::White(); return true;
            case 'yellow': $color = SKColors::Yellow(); return true;
            default:
                $color = new SKColor(0, 0, 0, 0);
                return false;
        }
    }
}
