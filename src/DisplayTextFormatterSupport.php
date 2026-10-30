<?php

namespace Aspose\Cells_FOSS;

final class DisplayTextFormatterSupport
{
    private function __construct()
    {
    }

    public static function SelectNumericSection($sections, $numericValue, &$useAbsoluteValue)
    {
        $useAbsoluteValue = false;
        $numericSectionCount = count($sections) < 3 ? count($sections) : 3;
        if ($numericSectionCount === 0) {
            return null;
        }

        $hasCondition = false;
        $fallbackSection = null;
        for ($index = 0; $index < $numericSectionCount; $index++) {
            $section = $sections[$index];
            if ($section->HasCondition) {
                $hasCondition = true;
                if (self::EvaluateCondition($section->ConditionOperator, $numericValue, $section->ConditionValue)) {
                    $useAbsoluteValue = self::ShouldUseAbsoluteValue($section->Raw, $numericValue);
                    return $section;
                }
            } else if ($fallbackSection === null) {
                $fallbackSection = $section;
            }
        }

        if ($hasCondition && $fallbackSection !== null) {
            $useAbsoluteValue = self::ShouldUseAbsoluteValue($fallbackSection->Raw, $numericValue);
            return $fallbackSection;
        }

        if ($numericSectionCount === 1) {
            // A lone section formats the signed value itself.
            $useAbsoluteValue = false;
            return $sections[0];
        }

        if ($numericSectionCount === 2) {
            if ($numericValue < 0) {
                // A dedicated negative section receives the magnitude. Its own
                // literals (minus sign, parentheses, currency) supply the sign.
                $useAbsoluteValue = true;
                return $sections[1];
            }

            $useAbsoluteValue = false;
            return $sections[0];
        }

        if ($numericValue > 0) {
            $useAbsoluteValue = false;
            return $sections[0];
        }

        if ($numericValue < 0) {
            $useAbsoluteValue = true;
            return $sections[1];
        }

        $useAbsoluteValue = false;
        return $sections[2];
    }

    public static function SelectTextSection($sections)
    {
        if (count($sections) >= 4) {
            return $sections[3];
        }

        for ($index = 0; $index < count($sections); $index++) {
            $section = self::StripDirectiveBrackets($sections[$index]->Raw, false);
            if (strpos($section, '@') !== false) {
                return $sections[$index];
            }
        }

        return null;
    }

    public static function SelectDateTimeSection($sections)
    {
        $dateSectionCount = count($sections) < 3 ? count($sections) : 3;
        for ($index = 0; $index < $dateSectionCount; $index++) {
            if (!($sections[$index]->Raw === null || trim($sections[$index]->Raw) === '')) {
                return $sections[$index];
            }
        }

        return null;
    }

    public static function ParseSections($formatCode)
    {
        $rawSections = self::SplitSections($formatCode);
        $sections = array();
        for ($index = 0; $index < count($rawSections); $index++) {
            $rawSection = $rawSections[$index];
            $section = new DisplayFormatSectionInfo();
            $section->Raw = $rawSection;

            $conditionOperator = '';
            $conditionValue = 0;
            if (self::TryParseSectionCondition($rawSection, $conditionOperator, $conditionValue)) {
                $section->HasCondition = true;
                $section->ConditionOperator = $conditionOperator;
                $section->ConditionValue = $conditionValue;
            }

            $sections[] = $section;
        }

        return $sections;
    }

    public static function SplitSections($formatCode)
    {
        $sections = array();
        $builder = '';
        $inQuote = false;

        for ($index = 0; $index < strlen($formatCode); $index++) {
            $character = $formatCode[$index];
            if ($character === '"') {
                $builder .= $character;
                $inQuote = !$inQuote;
                continue;
            }

            if (!$inQuote) {
                if ($character === '\\' && $index + 1 < strlen($formatCode)) {
                    $builder .= $character;
                    $index++;
                    $builder .= $formatCode[$index];
                    continue;
                }

                if ($character === ';') {
                    $sections[] = $builder;
                    $builder = '';
                    continue;
                }
            }

            $builder .= $character;
        }

        $sections[] = $builder;
        return $sections;
    }

    public static function TryParseSectionCondition($section, &$conditionOperator, &$conditionValue)
    {
        $conditionOperator = '';
        $conditionValue = 0;
        $inQuote = false;

        for ($index = 0; $index < strlen($section); $index++) {
            $character = $section[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($inQuote || $character !== '[') {
                continue;
            }

            $endIndex = strpos($section, ']', $index + 1);
            if ($endIndex === false) {
                break;
            }

            $token = trim(substr($section, $index + 1, $endIndex - $index - 1));
            if (self::TryParseConditionToken($token, $conditionOperator, $conditionValue)) {
                return true;
            }

            $index = $endIndex;
        }

        return false;
    }

    public static function TryParseConditionToken($token, &$conditionOperator, &$conditionValue)
    {
        $conditionOperator = '';
        $conditionValue = 0;
        if ($token === null || trim($token) === '') {
            return false;
        }

        $operators = array('>=', '<=', '<>', '>', '<', '=');
        for ($index = 0; $index < count($operators); $index++) {
            $candidate = $operators[$index];
            if (str_starts_with($token, $candidate)) {
                $numberPart = trim(substr($token, strlen($candidate)));
                if (is_numeric($numberPart)) {
                    $conditionOperator = $candidate;
                    $conditionValue = (float) $numberPart;
                    return true;
                }
            }
        }

        return false;
    }

    public static function EvaluateCondition($conditionOperator, $numericValue, $conditionValue)
    {
        switch ($conditionOperator) {
            case '>':
                return $numericValue > $conditionValue;
            case '>=':
                return $numericValue >= $conditionValue;
            case '<':
                return $numericValue < $conditionValue;
            case '<=':
                return $numericValue <= $conditionValue;
            case '=':
                return abs($numericValue - $conditionValue) < 1E-12;
            case '<>':
                return abs($numericValue - $conditionValue) >= 1E-12;
            default:
                return false;
        }
    }

    public static function ShouldUseAbsoluteValue($section, $numericValue)
    {
        if ($numericValue >= 0) {
            return false;
        }

        $sanitizedSection = self::SanitizeNumericSection($section);
        return strpos($sanitizedSection, '-') === false;
    }

    public static function SanitizeNumericSection($section)
    {
        $withoutDirectives = self::StripDirectiveBrackets($section, false);
        $builder = '';
        $inQuote = false;

        for ($index = 0; $index < strlen($withoutDirectives); $index++) {
            $character = $withoutDirectives[$index];
            if ($character === '"') {
                $builder .= $character;
                $inQuote = !$inQuote;
                continue;
            }

            if (!$inQuote) {
                if ($character === '_') {
                    $index++;
                    continue;
                }

                if ($character === '*') {
                    $index++;
                    continue;
                }

                if ($character === '\\') {
                    if ($index + 1 < strlen($withoutDirectives)) {
                        $index++;
                        $builder .= $withoutDirectives[$index];
                    }

                    continue;
                }

                if ($character === '?') {
                    $builder .= '#';
                    continue;
                }

                if ($character === '[' || $character === ']') {
                    continue;
                }
            }

            $builder .= $character;
        }

        return $builder;
    }

    public static function StripDirectiveBrackets($section, $preserveElapsedTokens)
    {
        $builder = '';
        $inQuote = false;

        for ($index = 0; $index < strlen($section); $index++) {
            $character = $section[$index];
            if ($character === '"') {
                $builder .= $character;
                $inQuote = !$inQuote;
                continue;
            }

            if (!$inQuote && $character === '[') {
                $endIndex = strpos($section, ']', $index + 1);
                if ($endIndex === false) {
                    continue;
                }

                $token = substr($section, $index + 1, $endIndex - $index - 1);
                if ($preserveElapsedTokens && DisplayTextDateFormatSupport::IsElapsedToken($token)) {
                    $builder .= '[';
                    $builder .= $token;
                    $builder .= ']';
                }

                $index = $endIndex;
                continue;
            }

            $builder .= $character;
        }

        return $builder;
    }

    public static function ContainsNumericPlaceholder($pattern)
    {
        for ($index = 0; $index < strlen($pattern); $index++) {
            $character = $pattern[$index];
            if ($character === '0' || $character === '#' || $character === '?' || $character === '.' || $character === '%' || $character === 'E' || $character === 'e' || $character === '/') {
                return true;
            }
        }

        return false;
    }

    public static function ExpandSectionPattern($pattern, $valueText, $replaceTextPlaceholder)
    {
        $builder = '';
        $inQuote = false;

        for ($index = 0; $index < strlen($pattern); $index++) {
            $character = $pattern[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if (!$inQuote) {
                if ($character === '@' && $replaceTextPlaceholder) {
                    $builder .= $valueText;
                    continue;
                }

                if ($character === '_') {
                    $index++;
                    continue;
                }

                if ($character === '*') {
                    $index++;
                    continue;
                }

                if ($character === '\\') {
                    if ($index + 1 < strlen($pattern)) {
                        $index++;
                        $builder .= $pattern[$index];
                    }

                    continue;
                }
            }

            $builder .= $character;
        }

        return $builder;
    }

    public static function IsNumericValue($value)
    {
        if (is_int($value)) {
            return true;
        }

        if (is_float($value)) {
            return true;
        }

        return false;
    }

    public static function ConvertToDouble($value)
    {
        if (is_int($value)) return (float) $value;
        if (is_float($value)) return $value;
        throw new \RuntimeException('Value is not numeric.');
    }

    public static function GetAbsoluteNumericValue($value)
    {
        if (is_int($value)) return abs($value);
        if (is_float($value)) return abs($value);
        return $value;
    }
}
