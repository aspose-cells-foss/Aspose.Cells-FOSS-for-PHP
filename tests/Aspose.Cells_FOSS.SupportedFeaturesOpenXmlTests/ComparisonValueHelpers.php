<?php

namespace Aspose\Cells_FOSS\SupportedFeaturesOpenXmlTests;

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class ComparisonValueHelpers
{
    const DefaultColorHex = '00000000';
    const SpreadsheetNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';

    private static $builtInDateFormats = array(14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47);

    public static function BuiltInDateFormats()
    {
        return self::$builtInDateFormats;
    }

    public static function IsBuiltInDateFormat($numberFormatId)
    {
        return in_array($numberFormatId, self::$builtInDateFormats, true);
    }

    public static function AttributeText($element, $name)
    {
        if ($element === null) {
            return null;
        }

        $attribute = $element->Attribute($name);
        return $attribute === null ? null : (string) $attribute->Value;
    }

    public static function TryNormalizeCellReference($cellReference, &$normalized)
    {
        $normalized = '';
        if ($cellReference === null || trim($cellReference) === '') {
            return false;
        }

        $trimmed = strtoupper(trim($cellReference));
        $splitIndex = 0;
        while ($splitIndex < strlen($trimmed) && $trimmed[$splitIndex] >= 'A' && $trimmed[$splitIndex] <= 'Z') {
            $splitIndex++;
        }

        if ($splitIndex === 0 || $splitIndex === strlen($trimmed)) {
            return false;
        }

        for ($index = $splitIndex; $index < strlen($trimmed); $index++) {
            if ($trimmed[$index] < '0' || $trimmed[$index] > '9') {
                return false;
            }
        }

        $row = (int) substr($trimmed, $splitIndex);
        if ($row < 1) {
            return false;
        }

        $normalized = $trimmed;
        return true;
    }

    public static function BuildCellKey($sheetName, $cellReference)
    {
        return $sheetName . '!' . strtoupper($cellReference);
    }

    public static function NormalizeValue($value)
    {
        if ($value === null) {
            return '<null>';
        }

        if (is_string($value)) {
            return $value;
        }

        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }

        if ($value instanceof \DateTime) {
            $utc = clone $value;
            $utc->setTimezone(new \DateTimeZone('UTC'));
            return $utc->format('Y-m-d\TH:i:s.0000000');
        }

        if (is_int($value)) {
            return (string) $value;
        }

        if (is_float($value)) {
            return (string) $value;
        }

        if (is_object($value) && method_exists($value, '__toString')) {
            return (string) $value;
        }

        return (string) $value;
    }

    public static function NormalizeFormulaText($formulaBody)
    {
        if ($formulaBody === null || trim($formulaBody) === '') {
            return '';
        }

        $normalized = trim($formulaBody);
        return strpos($normalized, '=') === 0 ? $normalized : '=' . $normalized;
    }

    public static function NormalizeDouble($value)
    {
        $text = number_format((float) $value, 4, '.', '');
        $text = rtrim(rtrim($text, '0'), '.');
        if ($text === '' || $text === '-') {
            $text = '0';
        }

        return $text;
    }

    public static function NormalizeColor(Color $color)
    {
        if ($color->Equals(Color::Empty())) {
            return self::DefaultColorHex;
        }

        return sprintf('%02X%02X%02X%02X', $color->A, $color->R, $color->G, $color->B);
    }

    public static function NormalizeOpenXmlColor($colorElement)
    {
        if ($colorElement === null) {
            return self::DefaultColorHex;
        }

        $rgb = self::AttributeText($colorElement, 'rgb');
        if ($rgb === null || trim($rgb) === '') {
            return self::DefaultColorHex;
        }

        $rgb = strtoupper(trim($rgb));
        if (strlen($rgb) === 6) {
            $rgb = 'FF' . $rgb;
        }

        return strlen($rgb) === 8 ? $rgb : self::DefaultColorHex;
    }

    public static function ParseHorizontalAlignment($value)
    {
        switch ($value === null ? null : strtolower($value)) {
            case 'left':
                return 'Left';
            case 'center':
                return 'Center';
            case 'right':
                return 'Right';
            case 'fill':
                return 'Fill';
            case 'justify':
                return 'Justify';
            case 'centercontinuous':
                return 'CenterContinuous';
            case 'distributed':
                return 'Distributed';
            default:
                return 'General';
        }
    }

    public static function ParseVerticalAlignment($value)
    {
        switch ($value === null ? null : strtolower($value)) {
            case 'center':
                return 'Center';
            case 'top':
                return 'Top';
            case 'justify':
                return 'Justify';
            case 'distributed':
                return 'Distributed';
            default:
                return 'Bottom';
        }
    }

    public static function ParseBorderStyle($value)
    {
        switch ($value === null ? null : strtolower($value)) {
            case 'thin':
                return 'Thin';
            case 'medium':
                return 'Medium';
            case 'thick':
                return 'Thick';
            case 'dotted':
                return 'Dotted';
            case 'dashed':
                return 'Dashed';
            case 'double':
                return 'Double';
            case 'hair':
                return 'Hair';
            case 'mediumdashed':
                return 'MediumDashed';
            case 'dashdot':
                return 'DashDot';
            case 'mediumdashdot':
                return 'MediumDashDot';
            case 'dashdotdot':
                return 'DashDotDot';
            case 'mediumdashdotdot':
                return 'MediumDashDotDot';
            case 'slantdashdot':
                return 'SlantedDashDot';
            default:
                return 'None';
        }
    }

    public static function ParseFillPattern($value)
    {
        switch ($value === null ? null : strtolower($value)) {
            case 'solid':
                return 'Solid';
            case 'mediumgray':
                return 'MediumGray';
            case 'darkgray':
                return 'DarkGray';
            case 'gray125':
                return 'Gray125';
            case 'gray0625':
                return 'Gray0625';
            case 'darkhorizontal':
                return 'DarkHorizontal';
            case 'darkvertical':
                return 'DarkVertical';
            case 'darkdown':
                return 'DarkDown';
            case 'darkup':
                return 'DarkUp';
            case 'darkgrid':
                return 'DarkGrid';
            case 'darktrellis':
                return 'DarkTrellis';
            case 'lighthorizontal':
                return 'LightHorizontal';
            case 'lightvertical':
                return 'LightVertical';
            case 'lightdown':
                return 'LightDown';
            case 'lightup':
                return 'LightUp';
            case 'lightgrid':
                return 'LightGrid';
            case 'lighttrellis':
                return 'LightTrellis';
            default:
                return 'None';
        }
    }

    public static function ParseBoolValue($value)
    {
        if ($value === null) {
            return false;
        }

        return $value === '1' || strcasecmp($value, 'true') === 0;
    }

    public static function ParseOptionalBoolValue($value)
    {
        return $value === null ? null : self::ParseBoolValue($value);
    }

    public static function ParseBoolAttribute($attribute)
    {
        if ($attribute === null) {
            return false;
        }

        return self::ParseBoolValue((string) $attribute->Value);
    }

    public static function ParseOptionalBoolAttribute($attribute)
    {
        return $attribute === null ? null : self::ParseBoolAttribute($attribute);
    }

    public static function ParseIntAttribute($attribute)
    {
        if ($attribute === null) {
            return null;
        }

        return self::ParseIntValue((string) $attribute->Value);
    }

    public static function ParseIntValue($value)
    {
        if ($value === null) {
            return null;
        }

        if (preg_match('/^\s*[+-]?\d+\s*$/', $value) !== 1) {
            return null;
        }

        return (int) $value;
    }

    public static function ParseDoubleValue($value, $defaultValue)
    {
        return is_numeric($value) ? (float) $value : $defaultValue;
    }

    public static function ParseDoubleValueOrNull($value)
    {
        return is_numeric($value) ? (float) $value : null;
    }

    public static function TryParseNumber($rawValue, &$value, &$kind)
    {
        if ($rawValue === null) {
            $kind = 'Blank';
            $value = null;
            return false;
        }

        $hasExponent = strpos($rawValue, 'e') !== false || strpos($rawValue, 'E') !== false;
        if (!$hasExponent && is_numeric($rawValue)) {
            $kind = 'Number';
            if (strpos($rawValue, '.') === false && strpos($rawValue, ',') === false) {
                $numeric = (float) $rawValue;
                if ($numeric >= -2147483648 && $numeric <= 2147483647) {
                    $value = (int) $rawValue;
                    return true;
                }
            }

            $value = (float) $rawValue;
            return true;
        }

        if (is_numeric($rawValue)) {
            $kind = 'Number';
            $doubleValue = (float) $rawValue;
            if (fmod($doubleValue, 1.0) == 0.0 && $doubleValue >= -2147483648 && $doubleValue <= 2147483647) {
                $value = (int) $doubleValue;
                return true;
            }

            $value = $doubleValue;
            return true;
        }

        $kind = 'Blank';
        $value = null;
        return false;
    }

    public static function LooksLikeDateFormat($formatCode)
    {
        if ($formatCode === null || trim($formatCode) === '') {
            return false;
        }

        $builder = '';
        $inQuote = false;
        $inBracket = false;
        $length = strlen($formatCode);

        for ($index = 0; $index < $length; $index++) {
            $character = $formatCode[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($inQuote) {
                continue;
            }

            if ($character === '[') {
                $inBracket = true;
                continue;
            }

            if ($character === ']' && $inBracket) {
                $inBracket = false;
                continue;
            }

            if ($inBracket) {
                $builder .= strtolower($character);
                continue;
            }

            if ($character === '\\' || $character === '_') {
                $index++;
                continue;
            }

            if ($character === '*') {
                continue;
            }

            $builder .= strtolower($character);
        }

        return strpos($builder, 'yy') !== false
            || strpos($builder, 'dd') !== false
            || strpos($builder, 'mm') !== false
            || strpos($builder, 'm/') !== false
            || strpos($builder, '/m') !== false
            || strpos($builder, 'h') !== false
            || strpos($builder, 'ss') !== false;
    }

    public static function GetOrDefault($values, $index, $fallback)
    {
        return $index >= 0 && $index < count($values) ? $values[$index] : $fallback;
    }

    public static function EnumName($enumClass, $value)
    {
        static $cache = array();
        if ($value === null) {
            $value = 0;
        }

        if (!array_key_exists($enumClass, $cache)) {
            $reflection = new \ReflectionClass($enumClass);
            $cache[$enumClass] = $reflection->getConstants();
        }

        foreach ($cache[$enumClass] as $name => $constantValue) {
            if ($constantValue === $value) {
                return $name;
            }
        }

        return (string) $value;
    }
}
