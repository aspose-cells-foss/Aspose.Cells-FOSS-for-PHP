<?php

namespace Aspose\Cells_FOSS;

/**
 * Emulates .NET custom numeric format strings for native PHP int/float values.
 *
 * Native PHP numbers are not IFormattable, so the .NET call
 * value.ToString(sanitizedSection, sectionCulture) is replaced by
 * {@see DisplayTextNumberSupport::FormatNumber()}, which understands the same
 * placeholder syntax used by the sanitized number format sections.
 */
final class DisplayTextNumberSupport
{
    private function __construct()
    {
    }

    public static function FormatNumber($value, $format, $culture)
    {
        if ($culture === null) {
            $culture = CultureInfo::GetInvariantCulture();
        }

        $numberFormat = $culture->NumberFormat;
        $decimalSeparator = $numberFormat->NumberDecimalSeparator;
        $groupSeparator = $numberFormat->NumberGroupSeparator;
        $groupSizes = $numberFormat->NumberGroupSizes;
        $percentSymbol = $numberFormat->PercentSymbol;
        $negativeSign = $numberFormat->NegativeSign;

        $tokens = self::Tokenize($format);
        if (count($tokens) === 0) {
            return '';
        }

        $hasNumericPlaceholder = false;
        $hasPercent = false;
        $scientificIndex = -1;
        for ($index = 0; $index < count($tokens); $index++) {
            $kind = $tokens[$index]['kind'];
            if ($kind === 'zero' || $kind === 'hash' || $kind === 'any' || $kind === 'decimal') {
                $hasNumericPlaceholder = true;
            }

            if ($kind === 'percent') {
                $hasPercent = true;
            }

            if ($kind === 'scientific') {
                $hasNumericPlaceholder = true;
                $scientificIndex = $index;
            }
        }

        if (!$hasNumericPlaceholder) {
            return self::ExpandLiterals($tokens);
        }

        $absolute = abs((float) $value);
        if ($hasPercent) {
            $absolute *= 100.0;
        }

        if ($scientificIndex >= 0) {
            $numberText = self::FormatScientific($absolute, $tokens, $scientificIndex, $decimalSeparator);
        } else {
            $numberText = self::FormatPositional($absolute, $tokens, $groupSeparator, $groupSizes, $decimalSeparator);
        }

        $hasNegativeLiteral = self::ContainsNegativeLiteral($tokens);

        $builder = '';
        $numericEmitted = false;
        for ($index = 0; $index < count($tokens); $index++) {
            $token = $tokens[$index];
            $kind = $token['kind'];
            if ($kind === 'literal') {
                $builder .= $token['text'];
                continue;
            }

            if ($kind === 'percent') {
                $builder .= $percentSymbol;
                continue;
            }

            if ($kind === 'zero' || $kind === 'hash' || $kind === 'any' || $kind === 'decimal' || $kind === 'comma' || $kind === 'scientific') {
                if (!$numericEmitted) {
                    if ((float) $value < 0 && !$hasNegativeLiteral) {
                        $builder .= $negativeSign;
                    }

                    $builder .= $numberText;
                    $numericEmitted = true;
                }

                continue;
            }
        }

        return $builder;
    }

    private static function FormatPositional($absolute, $tokens, $groupSeparator, $groupSizes, $decimalSeparator)
    {
        $intPlaceholders = array();
        $decPlaceholders = array();
        $commas = array();
        $afterDecimal = false;

        for ($index = 0; $index < count($tokens); $index++) {
            $kind = $tokens[$index]['kind'];
            if ($kind === 'scientific') {
                break;
            }

            if ($kind === 'decimal') {
                $afterDecimal = true;
                continue;
            }

            if ($kind === 'zero' || $kind === 'hash' || $kind === 'any') {
                if ($afterDecimal) {
                    $decPlaceholders[] = $kind;
                } else {
                    $intPlaceholders[] = $kind;
                }

                continue;
            }

            if ($kind === 'comma' && !$afterDecimal) {
                $commas[] = count($intPlaceholders);
            }
        }

        $intTotal = count($intPlaceholders);
        $grouping = false;
        $scaling = 0;
        for ($index = 0; $index < count($commas); $index++) {
            if ($commas[$index] < $intTotal) {
                $grouping = true;
            } else {
                $scaling++;
            }
        }

        if ($scaling > 0) {
            $absolute /= pow(1000.0, (float) $scaling);
        }

        $decimalCount = count($decPlaceholders);
        $rounded = number_format($absolute, $decimalCount, '.', '');
        $parts = explode('.', $rounded);
        $intDigits = $parts[0];
        $decDigits = count($parts) > 1 ? $parts[1] : '';

        $requiredZeros = 0;
        for ($index = 0; $index < count($intPlaceholders); $index++) {
            if ($intPlaceholders[$index] === 'zero') {
                $requiredZeros++;
            }
        }

        if (strlen($intDigits) < $requiredZeros) {
            $intDigits = str_pad($intDigits, $requiredZeros, '0', STR_PAD_LEFT);
        }

        if ($intDigits === '0' && $requiredZeros === 0) {
            $intDigits = '';
        }

        if ($grouping && $intDigits !== '') {
            $intDigits = self::InsertGroupSeparators($intDigits, $groupSizes, $groupSeparator);
        }

        $displayDecimals = self::TrimOptionalDecimals($decDigits, $decPlaceholders);

        $result = $intDigits;
        if ($displayDecimals !== '') {
            $result .= $decimalSeparator . $displayDecimals;
        }

        return $result;
    }

    private static function FormatScientific($absolute, $tokens, $scientificIndex, $decimalSeparator)
    {
        $scientific = $tokens[$scientificIndex];
        $exponentMark = $scientific['mark'];
        $exponentSign = $scientific['sign'];
        $exponentDigits = $scientific['digits'];

        $intPlaceholders = array();
        $decPlaceholders = array();
        $afterDecimal = false;
        for ($index = 0; $index < $scientificIndex; $index++) {
            $kind = $tokens[$index]['kind'];
            if ($kind === 'decimal') {
                $afterDecimal = true;
                continue;
            }

            if ($kind === 'zero' || $kind === 'hash' || $kind === 'any') {
                if ($afterDecimal) {
                    $decPlaceholders[] = $kind;
                } else {
                    $intPlaceholders[] = $kind;
                }
            }
        }

        $decimalCount = count($decPlaceholders);

        $exponent = 0;
        $mantissa = $absolute;
        if ($absolute != 0.0) {
            $exponent = (int) floor(log10($absolute));
            $mantissa = $absolute / pow(10.0, (float) $exponent);
        }

        $mantissa = round($mantissa, $decimalCount, PHP_ROUND_HALF_UP);
        if ($mantissa >= 10.0) {
            $mantissa = round($mantissa / 10.0, $decimalCount, PHP_ROUND_HALF_UP);
            $exponent++;
        }

        $rounded = number_format($mantissa, $decimalCount, '.', '');
        $parts = explode('.', $rounded);
        $intDigits = $parts[0];
        $decDigits = count($parts) > 1 ? $parts[1] : '';

        $requiredZeros = 0;
        for ($index = 0; $index < count($intPlaceholders); $index++) {
            if ($intPlaceholders[$index] === 'zero') {
                $requiredZeros++;
            }
        }

        if (strlen($intDigits) < $requiredZeros) {
            $intDigits = str_pad($intDigits, $requiredZeros, '0', STR_PAD_LEFT);
        }

        if ($intDigits === '0' && $requiredZeros === 0) {
            $intDigits = '';
        }

        $displayDecimals = self::TrimOptionalDecimals($decDigits, $decPlaceholders);

        $exponentValue = $exponent;
        $exponentText = '';
        if ($exponentValue < 0) {
            $exponentText = '-';
            $exponentValue = -$exponentValue;
        } else if ($exponentSign === '+') {
            $exponentText = '+';
        }

        $exponentText .= str_pad((string) $exponentValue, $exponentDigits, '0', STR_PAD_LEFT);

        $result = $intDigits;
        if ($displayDecimals !== '') {
            $result .= $decimalSeparator . $displayDecimals;
        }

        $result .= $exponentMark . $exponentText;
        return $result;
    }

    private static function TrimOptionalDecimals($decDigits, $decPlaceholders)
    {
        $lastRequired = -1;
        for ($index = 0; $index < count($decPlaceholders); $index++) {
            if ($decPlaceholders[$index] === 'zero' || $decDigits[$index] !== '0') {
                $lastRequired = $index;
            }
        }

        if ($lastRequired < 0) {
            return '';
        }

        return substr($decDigits, 0, $lastRequired + 1);
    }

    private static function InsertGroupSeparators($digits, $groupSizes, $separator)
    {
        if ($groupSizes === null || count($groupSizes) === 0) {
            return $digits;
        }

        $result = '';
        $index = strlen($digits) - 1;
        $groupIndex = 0;
        $countInGroup = 0;
        $currentSize = (int) $groupSizes[0];
        if ($currentSize <= 0) {
            return $digits;
        }

        while ($index >= 0) {
            if ($countInGroup === $currentSize) {
                $result = $separator . $result;
                $countInGroup = 0;
                if ($groupIndex < count($groupSizes) - 1) {
                    $groupIndex++;
                    $currentSize = (int) $groupSizes[$groupIndex];
                    if ($currentSize <= 0) {
                        $currentSize = 3;
                    }
                }
            }

            $result = $digits[$index] . $result;
            $index--;
            $countInGroup++;
        }

        return $result;
    }

    private static function ContainsNegativeLiteral($tokens)
    {
        for ($index = 0; $index < count($tokens); $index++) {
            $token = $tokens[$index];
            if ($token['kind'] === 'literal' && strpos($token['text'], '-') !== false) {
                return true;
            }
        }

        return false;
    }

    private static function ExpandLiterals($tokens)
    {
        $builder = '';
        for ($index = 0; $index < count($tokens); $index++) {
            $token = $tokens[$index];
            if ($token['kind'] === 'literal') {
                $builder .= $token['text'];
            }
        }

        return $builder;
    }

    private static function Tokenize($format)
    {
        $tokens = array();
        $inQuote = false;
        $length = strlen($format);

        for ($index = 0; $index < $length; $index++) {
            $character = $format[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($character === '\\') {
                if ($index + 1 < $length) {
                    $index++;
                    $tokens[] = array('kind' => 'literal', 'text' => $format[$index]);
                }

                continue;
            }

            if ($inQuote) {
                $tokens[] = array('kind' => 'literal', 'text' => $character);
                continue;
            }

            if ($character === '0') {
                $tokens[] = array('kind' => 'zero');
                continue;
            }

            if ($character === '#') {
                $tokens[] = array('kind' => 'hash');
                continue;
            }

            if ($character === '?') {
                $tokens[] = array('kind' => 'any');
                continue;
            }

            if ($character === '.') {
                $tokens[] = array('kind' => 'decimal');
                continue;
            }

            if ($character === ',') {
                $tokens[] = array('kind' => 'comma');
                continue;
            }

            if ($character === '%') {
                $tokens[] = array('kind' => 'percent');
                continue;
            }

            if ($character === 'E' || $character === 'e') {
                $scientific = self::TryReadScientific($format, $index);
                if ($scientific !== null) {
                    $tokens[] = $scientific['token'];
                    $index = $scientific['end'] - 1;
                    continue;
                }
            }

            $tokens[] = array('kind' => 'literal', 'text' => $character);
        }

        return $tokens;
    }

    private static function TryReadScientific($format, $startIndex)
    {
        $length = strlen($format);
        $mark = $format[$startIndex];
        $index = $startIndex + 1;
        $sign = '';
        if ($index < $length && ($format[$index] === '+' || $format[$index] === '-')) {
            $sign = $format[$index];
            $index++;
        }

        $digits = 0;
        while ($index < $length && ($format[$index] === '0' || $format[$index] === '#')) {
            $digits++;
            $index++;
        }

        if ($digits === 0) {
            return null;
        }

        return array(
            'token' => array('kind' => 'scientific', 'mark' => $mark, 'sign' => $sign, 'digits' => $digits),
            'end' => $index,
        );
    }
}
