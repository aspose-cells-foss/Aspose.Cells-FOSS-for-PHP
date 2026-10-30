<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\StyleValue;

final class DisplayTextFormatter
{
    private function __construct()
    {
    }

    public static function FormatStringValue($value)
    {
        if ($value === null) {
            return '';
        }

        if (is_string($value)) {
            return $value;
        }

        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }

        if ($value instanceof \DateTime) {
            return self::FormatRawDateTimeValue($value);
        }

        if (is_int($value)) {
            return (string) $value;
        }

        if (is_float($value)) {
            return (string) $value;
        }

        if ($value instanceof IFormattable) {
            return $value->ToString(null, CultureInfo::$InvariantCulture);
        }

        return (string) $value;
    }

    public static function FormatDisplayValue($value, StyleValue $style, $workbookCulture, $dateSystem = DateSystem::Windows1900)
    {
        if ($value === null) {
            return '';
        }

        if (is_string($value)) {
            return self::FormatTextValue($value, $style);
        }

        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }

        if ($value instanceof \DateTime) {
            return self::FormatDateTimeValue($value, $style, $workbookCulture, $dateSystem);
        }

        if (DisplayTextFormatterSupport::IsNumericValue($value)) {
            return self::FormatNumericValue($value, $style, $workbookCulture);
        }

        if ($value instanceof IFormattable) {
            return $value->ToString(null, $workbookCulture);
        }

        return (string) $value;
    }

    private static function FormatTextValue($value, StyleValue $style)
    {
        $formatCode = NumberFormat::ResolveFormatCode($style->NumberFormat->Number, $style->NumberFormat->Custom);
        if ($formatCode === null || trim($formatCode) === '' || $formatCode === 'General') {
            return $value;
        }

        // Excel text formats can have a dedicated fourth section; only that section
        // should shape the displayed text without mutating the underlying value.
        $sections = DisplayTextFormatterSupport::ParseSections($formatCode);
        $textSection = DisplayTextFormatterSupport::SelectTextSection($sections);
        if ($textSection === null) {
            return $value;
        }

        $pattern = DisplayTextFormatterSupport::StripDirectiveBrackets($textSection->Raw, false);
        $formatted = DisplayTextFormatterSupport::ExpandSectionPattern($pattern, $value, true);
        if (strlen($formatted) === 0) {
            return $value;
        }

        return $formatted;
    }

    private static function FormatNumericValue($value, StyleValue $style, $workbookCulture)
    {
        $formatCode = NumberFormat::ResolveFormatCode($style->NumberFormat->Number, $style->NumberFormat->Custom);
        if ($formatCode === null || trim($formatCode) === '' || $formatCode === 'General') {
            return self::FormatStringValue($value);
        }

        // Numeric formats can split positive, negative, zero, and text behavior into
        // separate sections, so the display path first selects the matching section.
        $sections = DisplayTextFormatterSupport::ParseSections($formatCode);
        if (count($sections) === 0) {
            return self::FormatStringValue($value);
        }

        $numericValue = DisplayTextFormatterSupport::ConvertToDouble($value);
        $useAbsoluteValue = false;
        $selectedSection = DisplayTextFormatterSupport::SelectNumericSection($sections, $numericValue, $useAbsoluteValue);
        if ($selectedSection === null || $selectedSection->Raw === null || trim($selectedSection->Raw) === '') {
            return self::FormatStringValue($value);
        }

        $fractionResult = '';
        // Fraction placeholders are not directly compatible with .NET numeric format
        // strings, so they are handled separately before standard formatting.
        if (self::TryFormatFraction($numericValue, $selectedSection->Raw, $useAbsoluteValue, $fractionResult)) {
            return $fractionResult;
        }

        $sectionCulture = null;
        $localizedSection = DisplayTextLocaleSupport::ApplyLocaleDirectives($selectedSection->Raw, $workbookCulture, $sectionCulture);
        if (strpos($selectedSection->Raw, '-401]') !== false
            && $sectionCulture->NumberFormat->NumberDecimalSeparator !== '.') {
            $sectionCulture = $sectionCulture->Clone();
            $sectionCulture->NumberFormat->NumberDecimalSeparator = '.';
        }

        $sanitizedSection = DisplayTextFormatterSupport::SanitizeNumericSection($localizedSection);
        if ($sanitizedSection === null || trim($sanitizedSection) === '') {
            return self::FormatStringValue($value);
        }

        if (!DisplayTextFormatterSupport::ContainsNumericPlaceholder($sanitizedSection)) {
            $literal = DisplayTextFormatterSupport::ExpandSectionPattern($sanitizedSection, '', false);
            if (strlen($literal) > 0) {
                return $literal;
            }

            return self::FormatStringValue($value);
        }

        try {
            $formattedValue = $value;
            if ($useAbsoluteValue && $numericValue < 0) {
                $formattedValue = DisplayTextFormatterSupport::GetAbsoluteNumericValue($value);
            }

            if ($formattedValue instanceof IFormattable) {
                return $formattedValue->ToString($sanitizedSection, $sectionCulture);
            }

            return DisplayTextNumberSupport::FormatNumber($formattedValue, $sanitizedSection, $sectionCulture);
        } catch (\Exception $exception) {
            return self::FormatStringValue($value);
        }

        return self::FormatStringValue($value);
    }

    private static function FormatDateTimeValue($value, StyleValue $style, $workbookCulture, $dateSystem)
    {
        $formatCode = NumberFormat::ResolveFormatCode($style->NumberFormat->Number, $style->NumberFormat->Custom);
        if ($formatCode === null || trim($formatCode) === '' || $formatCode === 'General') {
            return self::FormatRawDateTimeValue($value);
        }

        // Date/time display uses the first section that still carries date tokens
        // after locale and directive cleanup.
        $sections = DisplayTextFormatterSupport::ParseSections($formatCode);
        $section = DisplayTextFormatterSupport::SelectDateTimeSection($sections);
        if ($section === null || $section->Raw === null || trim($section->Raw) === '') {
            return self::FormatRawDateTimeValue($value);
        }

        $sectionCulture = null;
        $localizedSection = DisplayTextLocaleSupport::ApplyLocaleDirectives($section->Raw, $workbookCulture, $sectionCulture);
        $sectionFormat = DisplayTextFormatterSupport::StripDirectiveBrackets($localizedSection, true);
        // Elapsed-time formats are duration-style output, not calendar rendering.
        if (DisplayTextDateFormatSupport::ContainsElapsedTimeToken($sectionFormat)) {
            $serial = DateSerialConverter::ToSerial($value, $dateSystem);
            return DisplayTextDateFormatSupport::FormatElapsedTimeValue($serial, $sectionFormat, $sectionCulture);
        }

        if ($sectionFormat === null || trim($sectionFormat) === '') {
            return self::FormatRawDateTimeValue($value);
        }

        try {
            $result = DisplayTextDateFormatSupport::FormatDateTimeValue($value, $sectionFormat, $sectionCulture);
            if (stripos($section->Raw, '[DBNum1]') !== false) {
                return self::FormatDbNumOne($result);
            }

            return $result;
        } catch (\Exception $exception) {
            return self::FormatRawDateTimeValue($value);
        }
    }

    private static function FormatDbNumOne($text)
    {
        $digits = array('零', '一', '二', '三', '四', '五', '六', '七', '八', '九');
        $yearChar = '年';
        $yearLength = strlen($yearChar);
        $builder = '';
        for ($index = 0; $index < strlen($text); $index++) {
            if ($text[$index] < '0' || $text[$index] > '9') {
                $builder .= $text[$index];
                continue;
            }

            $end = $index + 1;
            while ($end < strlen($text) && $text[$end] >= '0' && $text[$end] <= '9') {
                $end++;
            }

            if ($end < strlen($text) && substr($text, $end, $yearLength) === $yearChar) {
                for ($digitIndex = $index; $digitIndex < $end; $digitIndex++) {
                    $builder .= $digits[(int) $text[$digitIndex]];
                }
            } else {
                $digitsText = substr($text, $index, $end - $index);
                $number = ctype_digit($digitsText) ? (int) $digitsText : -1;
                if ($number >= 0 && $number < 100) {
                    if ($number >= 10) {
                        if ($number >= 20) {
                            $builder .= $digits[intdiv($number, 10)];
                        }

                        $builder .= '十';
                    }

                    if ($number % 10 !== 0 || $number === 0) {
                        $builder .= $digits[$number % 10];
                    }
                } else {
                    $builder .= $digitsText;
                }
            }

            $index = $end - 1;
        }

        return $builder;
    }

    private static function FormatRawDateTimeValue($value)
    {
        if ($value->format('His') === '000000') {
            return $value->format('n/j/Y');
        }

        return $value->format('n/j/Y G:i');
    }

    private static function TryFormatFraction($numericValue, $section, $useAbsoluteValue, &$result)
    {
        $result = '';
        $sanitizedSection = DisplayTextFormatterSupport::SanitizeNumericSection($section);
        if (strpos($sanitizedSection, '/') === false) {
            return false;
        }

        $slashIndex = strpos($sanitizedSection, '/');
        if ($slashIndex <= 0) {
            return false;
        }

        $fixedEnd = $slashIndex + 1;
        while ($fixedEnd < strlen($sanitizedSection) && ctype_digit($sanitizedSection[$fixedEnd])) {
            $fixedEnd++;
        }

        if ($fixedEnd > $slashIndex + 1) {
            $fixedDenominatorText = substr($sanitizedSection, $slashIndex + 1, $fixedEnd - $slashIndex - 1);
            if (!ctype_digit($fixedDenominatorText)) {
                return false;
            }

            $fixedDenominator = (int) $fixedDenominatorText;
            if ($fixedDenominator <= 0) {
                return false;
            }

            $absolute = abs($numericValue);
            $whole = (int) floor($absolute);
            $numerator = (int) round(($absolute - $whole) * $fixedDenominator, 0, PHP_ROUND_HALF_UP);
            if ($numerator >= $fixedDenominator) {
                $whole++;
                $numerator = 0;
            }

            $sign = $numericValue < 0 && !$useAbsoluteValue ? '-' : '';
            $result = $numerator === 0
                ? $sign . (string) $whole
                : $sign . ($whole === 0 ? '' : (string) $whole . ' ')
                    . (string) $numerator . '/' . (string) $fixedDenominator;
            return true;
        }

        $denominatorDigits = 0;
        for ($index = $slashIndex + 1; $index < strlen($sanitizedSection); $index++) {
            $character = $sanitizedSection[$index];
            if ($character === '#' || $character === '0') {
                $denominatorDigits++;
                continue;
            }

            if (ctype_space($character)) {
                continue;
            }

            break;
        }

        if ($denominatorDigits <= 0) {
            return false;
        }

        $absoluteValue = abs($numericValue);
        $wholePart = (int) floor($absoluteValue);
        $fractionalPart = $absoluteValue - $wholePart;
        if ($fractionalPart < 1E-12) {
            $result = self::FormatWholeFractionResult($wholePart, $useAbsoluteValue, $numericValue);
            return true;
        }

        $maxDenominator = 1;
        for ($index = 0; $index < $denominatorDigits; $index++) {
            $maxDenominator *= 10;
        }

        // Match Excel's visible fraction behavior by finding the closest rational
        // value that fits within the placeholder width from the format section.
        $maxDenominator -= 1;
        $bestNumerator = 0;
        $bestDenominator = 1;
        $bestError = PHP_FLOAT_MAX;

        for ($denominator = 1; $denominator <= $maxDenominator; $denominator++) {
            $numerator = (int) round($fractionalPart * $denominator, 0, PHP_ROUND_HALF_UP);
            if ($numerator === 0) {
                continue;
            }

            if ($numerator > $denominator) {
                $numerator = $denominator;
            }

            $candidate = (float) $numerator / $denominator;
            $error = abs($fractionalPart - $candidate);
            if ($error < $bestError) {
                $bestError = $error;
                $bestNumerator = $numerator;
                $bestDenominator = $denominator;
            }
        }

        if ($bestNumerator === 0) {
            $result = self::FormatWholeFractionResult($wholePart, $useAbsoluteValue, $numericValue);
            return true;
        }

        $greatestCommonDivisor = self::GreatestCommonDivisor($bestNumerator, $bestDenominator);
        $bestNumerator = intdiv($bestNumerator, $greatestCommonDivisor);
        $bestDenominator = intdiv($bestDenominator, $greatestCommonDivisor);

        if ($bestNumerator === $bestDenominator) {
            $wholePart++;
            $bestNumerator = 0;
        }

        $prefix = '';
        if (!$useAbsoluteValue && $numericValue < 0) {
            $prefix = '-';
        }

        if ($bestNumerator === 0) {
            $result = $prefix . (string) $wholePart;
            return true;
        }

        if ($wholePart === 0) {
            $result = $prefix . (string) $bestNumerator . '/' . (string) $bestDenominator;
            return true;
        }

        $result = $prefix . (string) $wholePart . ' ' . (string) $bestNumerator . '/' . (string) $bestDenominator;
        return true;
    }

    private static function FormatWholeFractionResult($wholePart, $useAbsoluteValue, $numericValue)
    {
        if (!$useAbsoluteValue && $numericValue < 0) {
            return '-' . (string) $wholePart;
        }

        return (string) $wholePart;
    }

    private static function GreatestCommonDivisor($left, $right)
    {
        $first = abs($left);
        $second = abs($right);
        while ($second !== 0) {
            $remainder = $first % $second;
            $first = $second;
            $second = $remainder;
        }

        if ($first === 0) {
            return 1;
        }

        return $first;
    }
}
