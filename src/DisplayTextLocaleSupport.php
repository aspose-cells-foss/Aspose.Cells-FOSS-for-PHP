<?php

namespace Aspose\Cells_FOSS;

final class DisplayTextLocaleSupport
{
    private function __construct()
    {
    }

    public static function ApplyLocaleDirectives($section, $fallbackCulture, &$sectionCulture)
    {
        $sectionCulture = $fallbackCulture;
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
                if ($endIndex !== false && $endIndex > $index) {
                    $token = substr($section, $index + 1, $endIndex - $index - 1);
                    $replacement = '';
                    $resolvedCulture = null;
                    if (self::TryResolveLocaleDirective($token, $sectionCulture, $replacement, $resolvedCulture)) {
                        $builder .= $replacement;
                        $sectionCulture = $resolvedCulture;
                        if (strcasecmp($token, '$-F800') === 0
                            || strcasecmp($token, '$-F400') === 0) {
                            // Excel's system long-date/time directives replace the
                            // stored fallback pattern rather than prefixing it.
                            return $builder;
                        }

                        $index = $endIndex;
                        continue;
                    }
                }
            }

            $builder .= $character;
        }

        return $builder;
    }

    private static function TryResolveLocaleDirective($token, $fallbackCulture, &$replacement, &$resolvedCulture)
    {
        $replacement = '';
        $resolvedCulture = $fallbackCulture;
        if ($token === null || trim($token) === '' || $token[0] !== '$') {
            return false;
        }

        $dashIndex = strrpos($token, '-');
        if ($dashIndex === false) {
            $replacement = self::QuoteLiteral(substr($token, 1));
            return true;
        }

        if ($dashIndex === 0 || $dashIndex >= strlen($token) - 1) {
            return false;
        }

        $symbol = substr($token, 1, $dashIndex - 1);
        $localeCode = substr($token, $dashIndex + 1);
        $resolvedCulture = self::ResolveCulture($localeCode, $fallbackCulture);

        if (strcasecmp($localeCode, 'F800') === 0) {
            $replacement = str_starts_with($fallbackCulture->Name, 'zh')
                ? 'yyyy"年"m"月"d"日"'
                : $fallbackCulture->DateTimeFormat->LongDatePattern;
            $resolvedCulture = $fallbackCulture;
            return true;
        }

        if (strcasecmp($localeCode, 'F400') === 0) {
            $replacement = str_starts_with($fallbackCulture->Name, 'zh')
                ? 'H:mm:ss'
                : $fallbackCulture->DateTimeFormat->LongTimePattern;
            $resolvedCulture = $fallbackCulture;
            return true;
        }

        if (!($symbol === null || $symbol === '')) {
            $replacement = self::QuoteLiteral($symbol);
        }

        return true;
    }

    private static function ResolveCulture($localeCode, $fallbackCulture)
    {
        if ($localeCode === null || trim($localeCode) === '') {
            return $fallbackCulture;
        }

        $trimmed = trim($localeCode);
        if (!ctype_xdigit($trimmed)) {
            return $fallbackCulture;
        }

        $lcid = (int) hexdec($trimmed);

        try {
            return CultureInfo::GetCultureInfo($lcid);
        } catch (\Exception $exception) {
            return $fallbackCulture;
        }
    }

    private static function QuoteLiteral($value)
    {
        $builder = '';
        $builder .= '"';
        for ($index = 0; $index < strlen($value); $index++) {
            if ($value[$index] === '"') {
                $builder .= '"';
            }

            $builder .= $value[$index];
        }

        $builder .= '"';
        return $builder;
    }
}
