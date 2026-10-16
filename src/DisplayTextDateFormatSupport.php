<?php

namespace Aspose\Cells_FOSS;

final class DisplayTextDateFormatSupport
{
    private function __construct()
    {
    }

    public static function FormatDateTimeValue($value, $formatCode, $culture)
    {
        $builder = '';
        $hasAmPm = stripos($formatCode, 'AM/PM') !== false
            || stripos($formatCode, 'A/P') !== false;
        $inQuote = false;

        for ($index = 0; $index < strlen($formatCode); $index++) {
            if (self::MatchesToken($formatCode, $index, 'AM/PM')) {
                $builder .= self::GetAmPmDesignator($value, $culture, false);
                $index += 4;
                continue;
            }

            if (self::MatchesToken($formatCode, $index, 'A/P')) {
                $builder .= self::GetAmPmDesignator($value, $culture, true);
                $index += 2;
                continue;
            }

            $character = $formatCode[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($inQuote) {
                $builder .= $character;
                continue;
            }

            if ($character === '\\') {
                if ($index + 1 < strlen($formatCode)) {
                    $index++;
                    $builder .= $formatCode[$index];
                }

                continue;
            }

            if ($character === '_' || $character === '*') {
                $index++;
                continue;
            }

            if ($character === '.' && self::TryAppendFractionalSeconds($value, $formatCode, $culture, $index, $builder)) {
                continue;
            }

            if ($character === 'y' || $character === 'Y') {
                $count = self::CountRepeated($formatCode, $index, $character);
                self::AppendYear($value, $culture, $count, $builder);
                $index += $count - 1;
                continue;
            }

            if ($character === 'd' || $character === 'D') {
                $count = self::CountRepeated($formatCode, $index, $character);
                self::AppendDay($value, $culture, $count, $builder);
                $index += $count - 1;
                continue;
            }

            if ($character === 'h' || $character === 'H') {
                $count = self::CountRepeated($formatCode, $index, $character);
                self::AppendHour($value, $culture, $count, $hasAmPm, $builder);
                $index += $count - 1;
                continue;
            }

            if ($character === 's' || $character === 'S') {
                $count = self::CountRepeated($formatCode, $index, $character);
                self::AppendSecond($value, $culture, $count, $builder);
                $index += $count - 1;
                continue;
            }

            if ($character === 'm' || $character === 'M') {
                $count = self::CountRepeated($formatCode, $index, $character);
                if (self::IsMinuteContext($formatCode, $index, $count)) {
                    self::AppendMinute($value, $culture, $count, $builder);
                } else {
                    self::AppendMonth($value, $culture, $count, $builder);
                }

                $index += $count - 1;
                continue;
            }

            $builder .= $character;
        }

        return $builder;
    }

    public static function MatchesToken($formatCode, $startIndex, $token)
    {
        if ($startIndex + strlen($token) > strlen($formatCode)) {
            return false;
        }

        return strcasecmp(substr($formatCode, $startIndex, strlen($token)), $token) === 0;
    }

    public static function CountRepeated($formatCode, $startIndex, $token)
    {
        $count = 1;
        for ($index = $startIndex + 1; $index < strlen($formatCode); $index++) {
            if (strtolower($formatCode[$index]) !== strtolower($token)) {
                break;
            }

            $count++;
        }

        return $count;
    }

    public static function IsMinuteContext($formatCode, $startIndex, $count)
    {
        $previous = self::FindNeighborToken($formatCode, $startIndex - 1, -1);
        $next = self::FindNeighborToken($formatCode, $startIndex + $count, 1);
        if ($previous === ':' || $next === ':') {
            return true;
        }

        if ($previous === 'h' || $previous === 'H' || $previous === 's' || $previous === 'S') {
            return true;
        }

        if ($next === 'h' || $next === 'H' || $next === 's' || $next === 'S') {
            return true;
        }

        return false;
    }

    public static function FindNeighborToken($formatCode, $startIndex, $direction)
    {
        $inQuote = false;
        if ($direction < 0) {
            for ($index = $startIndex; $index >= 0; $index--) {
                $character = $formatCode[$index];
                if ($character === '"') {
                    $inQuote = !$inQuote;
                    continue;
                }

                if ($inQuote) {
                    continue;
                }

                if ($character === '\\' || $character === '_' || $character === '*') {
                    $index--;
                    continue;
                }

                if (ctype_space($character)) {
                    continue;
                }

                return $character;
            }
        } else {
            for ($index = $startIndex; $index < strlen($formatCode); $index++) {
                $character = $formatCode[$index];
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

                if (ctype_space($character)) {
                    continue;
                }

                return $character;
            }
        }

        return "\0";
    }

    public static function IsElapsedToken($token)
    {
        $normalized = strtolower(trim($token));
        return $normalized === 'h' || $normalized === 'hh' || $normalized === 'm' || $normalized === 'mm' || $normalized === 's' || $normalized === 'ss';
    }

    public static function ContainsElapsedTimeToken($formatCode)
    {
        for ($index = 0; $index < strlen($formatCode); $index++) {
            if ($formatCode[$index] !== '[') {
                continue;
            }

            $endIndex = strpos($formatCode, ']', $index + 1);
            if ($endIndex === false) {
                continue;
            }

            $token = substr($formatCode, $index + 1, $endIndex - $index - 1);
            if (self::IsElapsedToken($token)) {
                return true;
            }
        }

        return false;
    }

    public static function FormatElapsedTimeValue($time, $formatCode, $culture)
    {
        $hours = (int) $time->format('G');
        $minutes = (int) $time->format('i');
        $seconds = (int) $time->format('s');
        $milliseconds = (int) $time->format('v');
        $totalSeconds = $hours * 3600 + $minutes * 60 + $seconds;
        $totalHours = (int) floor($totalSeconds / 3600);
        $totalMinutes = (int) floor($totalSeconds / 60);
        $builder = '';
        $inQuote = false;

        for ($index = 0; $index < strlen($formatCode); $index++) {
            $character = $formatCode[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($inQuote) {
                $builder .= $character;
                continue;
            }

            if ($character === '[') {
                $endIndex = strpos($formatCode, ']', $index + 1);
                if ($endIndex !== false && $endIndex > $index) {
                    $token = strtolower(substr($formatCode, $index + 1, $endIndex - $index - 1));
                    if ($token === 'h') {
                        $builder .= (string) $totalHours;
                    } else if ($token === 'hh') {
                        $builder .= str_pad((string) $totalHours, 2, '0', STR_PAD_LEFT);
                    } else if ($token === 'm') {
                        $builder .= (string) $totalMinutes;
                    } else if ($token === 'mm') {
                        $builder .= str_pad((string) $totalMinutes, 2, '0', STR_PAD_LEFT);
                    } else if ($token === 's') {
                        $builder .= (string) $totalSeconds;
                    } else if ($token === 'ss') {
                        $builder .= str_pad((string) $totalSeconds, 2, '0', STR_PAD_LEFT);
                    } else {
                        $builder .= '[';
                        $builder .= $token;
                        $builder .= ']';
                    }

                    $index = $endIndex;
                    continue;
                }
            }

            if ($character === 'h' || $character === 'H') {
                $count = self::CountRepeated($formatCode, $index, $character);
                if ($count === 1) {
                    $builder .= (string) $hours;
                } else {
                    $builder .= str_pad((string) $hours, 2, '0', STR_PAD_LEFT);
                }

                $index += $count - 1;
                continue;
            }

            if ($character === 'm' || $character === 'M') {
                $count = self::CountRepeated($formatCode, $index, $character);
                if ($count === 1) {
                    $builder .= (string) $minutes;
                } else {
                    $builder .= str_pad((string) $minutes, 2, '0', STR_PAD_LEFT);
                }

                $index += $count - 1;
                continue;
            }

            if ($character === 's' || $character === 'S') {
                $count = self::CountRepeated($formatCode, $index, $character);
                if ($count === 1) {
                    $builder .= (string) $seconds;
                } else {
                    $builder .= str_pad((string) $seconds, 2, '0', STR_PAD_LEFT);
                }

                $index += $count - 1;
                continue;
            }

            if ($character === '.' && $index + 1 < strlen($formatCode) && $formatCode[$index + 1] === '0') {
                $zeroCount = self::CountRepeated($formatCode, $index + 1, '0');
                $builder .= $culture->NumberFormat->NumberDecimalSeparator;
                self::AppendFractionDigits($milliseconds, $zeroCount, $builder);
                $index += $zeroCount;
                continue;
            }

            if ($character === '\\') {
                if ($index + 1 < strlen($formatCode)) {
                    $index++;
                    $builder .= $formatCode[$index];
                }

                continue;
            }

            if ($character === '_' || $character === '*') {
                $index++;
                continue;
            }

            $builder .= $character;
        }

        return $builder;
    }

    private static function AppendYear($value, $culture, $count, &$builder)
    {
        $year = (int) $value->format('Y');
        if ($count <= 1) {
            $builder .= (string) ($year % 100);
            return;
        }

        if ($count === 2) {
            $builder .= str_pad((string) ($year % 100), 2, '0', STR_PAD_LEFT);
            return;
        }

        $builder .= str_pad((string) $year, $count, '0', STR_PAD_LEFT);
    }

    private static function AppendDay($value, $culture, $count, &$builder)
    {
        if ($count === 1) {
            $builder .= (string) ((int) $value->format('j'));
            return;
        }

        if ($count === 2) {
            $builder .= str_pad((string) ((int) $value->format('j')), 2, '0', STR_PAD_LEFT);
            return;
        }

        if ($count === 3) {
            $builder .= $culture->DateTimeFormat->GetAbbreviatedDayName((int) $value->format('w'));
            return;
        }

        $builder .= $culture->DateTimeFormat->GetDayName((int) $value->format('w'));
    }

    private static function AppendMonth($value, $culture, $count, &$builder)
    {
        if ($count === 1) {
            $builder .= (string) ((int) $value->format('n'));
            return;
        }

        if ($count === 2) {
            $builder .= str_pad((string) ((int) $value->format('n')), 2, '0', STR_PAD_LEFT);
            return;
        }

        if ($count === 3) {
            $builder .= $culture->DateTimeFormat->GetAbbreviatedMonthName((int) $value->format('n'));
            return;
        }

        $monthName = $culture->DateTimeFormat->GetMonthName((int) $value->format('n'));
        if ($count === 4) {
            $builder .= $monthName;
            return;
        }

        if (strlen($monthName) === 0) {
            $builder .= (string) ((int) $value->format('n'));
            return;
        }

        $builder .= $monthName[0];
    }

    private static function AppendHour($value, $culture, $count, $hasAmPm, &$builder)
    {
        $hour = (int) $value->format('G');
        if ($hasAmPm) {
            $hour %= 12;
            if ($hour === 0) {
                $hour = 12;
            }
        }

        if ($count === 1) {
            $builder .= (string) $hour;
            return;
        }

        $builder .= str_pad((string) $hour, 2, '0', STR_PAD_LEFT);
    }

    private static function AppendMinute($value, $culture, $count, &$builder)
    {
        if ($count === 1) {
            $builder .= (string) ((int) $value->format('i'));
            return;
        }

        $builder .= str_pad((string) ((int) $value->format('i')), 2, '0', STR_PAD_LEFT);
    }

    private static function AppendSecond($value, $culture, $count, &$builder)
    {
        if ($count === 1) {
            $builder .= (string) ((int) $value->format('s'));
            return;
        }

        $builder .= str_pad((string) ((int) $value->format('s')), 2, '0', STR_PAD_LEFT);
    }

    private static function TryAppendFractionalSeconds($value, $formatCode, $culture, &$index, &$builder)
    {
        if ($index + 1 >= strlen($formatCode) || $formatCode[$index + 1] !== '0') {
            return false;
        }

        $previous = self::FindNeighborToken($formatCode, $index - 1, -1);
        if ($previous !== 's' && $previous !== 'S') {
            return false;
        }

        $zeroCount = self::CountRepeated($formatCode, $index + 1, '0');
        $builder .= $culture->NumberFormat->NumberDecimalSeparator;
        self::AppendFractionDigits((int) $value->format('v'), $zeroCount, $builder);
        $index += $zeroCount;
        return true;
    }

    private static function AppendFractionDigits($milliseconds, $zeroCount, &$builder)
    {
        $digits = str_pad((string) $milliseconds, 3, '0', STR_PAD_LEFT);
        if ($zeroCount <= 0) {
            return;
        }

        if ($zeroCount === 1) {
            $builder .= $digits[0];
            return;
        }

        if ($zeroCount === 2) {
            $builder .= $digits[0];
            $builder .= $digits[1];
            return;
        }

        $builder .= $digits;
        for ($index = 3; $index < $zeroCount; $index++) {
            $builder .= '0';
        }
    }

    private static function GetAmPmDesignator($value, $culture, $abbreviated)
    {
        $hour = (int) $value->format('G');
        $designator = $hour < 12 ? $culture->DateTimeFormat->AMDesignator : $culture->DateTimeFormat->PMDesignator;
        if ($designator === null || $designator === '') {
            $designator = $hour < 12 ? 'AM' : 'PM';
        }

        if (!$abbreviated) {
            return $designator;
        }

        return substr($designator, 0, 1);
    }
}
