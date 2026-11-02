<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class ConditionalFormatEvaluator
{
    private static $ArrowGreen = null;
    private static $ArrowAmber = null;
    private static $ArrowRed = null;
    private static $TrafficRed = null;
    private static $TrafficYellow = null;
    private static $TrafficGreen = null;
    private static $TrafficRedBorder = null;
    private static $TrafficYellowBorder = null;
    private static $TrafficGreenBorder = null;

    private $_effects = array();

    public function __construct($sheet, $context)
    {
        if ($sheet === null || $sheet->ConditionalFormattings === null) {
            return;
        }

        // Higher priority (lower number) wins for overlapping differential-format rules.
        $formattings = $sheet->ConditionalFormattings;
        foreach ($formattings as $formatting) {
            $conditions = $formatting->Conditions;
            usort($conditions, function ($a, $b) {
                return $a->Priority <=> $b->Priority;
            });

            foreach ($conditions as $condition) {
                foreach ($formatting->Areas as $area) {
                    $this->ApplyConditionToArea($sheet, $context, $condition, $area);
                }
            }
        }
    }

    private static function InitializeColors()
    {
        if (self::$ArrowGreen !== null) {
            return;
        }

        self::$ArrowGreen = new SKColor(0x63, 0xA5, 0x37);
        self::$ArrowAmber = new SKColor(0xE8, 0xA8, 0x38);
        self::$ArrowRed = new SKColor(0xDC, 0x3A, 0x2E);
        self::$TrafficRed = new SKColor(214, 85, 50);
        self::$TrafficYellow = new SKColor(234, 194, 130);
        self::$TrafficGreen = new SKColor(104, 164, 144);
        self::$TrafficRedBorder = new SKColor(161, 63, 36);
        self::$TrafficYellowBorder = new SKColor(167, 132, 51);
        self::$TrafficGreenBorder = new SKColor(58, 116, 97);
    }

    public function TryGet($row, $col, &$effect)
    {
        $key = self::Key($row, $col);
        if (array_key_exists($key, $this->_effects)) {
            $effect = $this->_effects[$key];
            return true;
        }

        return false;
    }

    private function EffectFor($row, $col)
    {
        $key = self::Key($row, $col);
        if (!array_key_exists($key, $this->_effects)) {
            $this->_effects[$key] = new CellEffect();
        }

        return $this->_effects[$key];
    }

    private function ApplyConditionToArea($sheet, $context, $condition, $area)
    {
        switch ($condition->Type) {
            case FormatConditionType::ColorScale:
                $this->ApplyColorScale($sheet, $context, $condition, $area);
                return;
            case FormatConditionType::DataBar:
                $this->ApplyDataBar($sheet, $context, $condition, $area);
                return;
            case FormatConditionType::IconSet:
                $this->ApplyIconSet($sheet, $condition, $area);
                return;
            case FormatConditionType::Expression:
                $this->ApplyExpression($sheet, $context, $condition, $area);
                return;
        }

        // Differential-format (dxf) rules: apply the fill/font only where the rule matches. A rule
        // with no dxf still "matches" but paints nothing (Excel behaviour for this file).
        $matcher = self::BuildMatcher($sheet, $context, $condition, $area);
        if ($matcher === null) {
            return;
        }

        for ($r = $area->StartRow; $r <= $area->EndRow; $r++) {
            for ($c = $area->StartColumn; $c <= $area->EndColumn; $c++) {
                $key = (new CellAddress($r, $c))->ToString();
                $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
                if (!($matcher)($record)) {
                    continue;
                }

                self::ApplyDifferentialStyle($context, $this->EffectFor($r, $c), $condition);
            }
        }
    }

    private function ApplyExpression($sheet, $context, $condition, $area)
    {
        $matcher = ConditionalFormatExpressionMatcher::TryCreate($sheet, $area, $condition->Formula1);
        if ($matcher === null) {
            return;
        }

        for ($row = $area->StartRow; $row <= $area->EndRow; $row++) {
            for ($column = $area->StartColumn; $column <= $area->EndColumn; $column++) {
                if ($matcher->Matches($row, $column)) {
                    self::ApplyDifferentialStyle($context, $this->EffectFor($row, $column), $condition);
                }
            }
        }
    }

    private static function ApplyDifferentialStyle($context, $effect, $condition)
    {
        $style = $condition->Style;
        if ($style === null) {
            return;
        }

        if (!$effect->HasFill && $style->Pattern !== FillPatternKind::None) {
            $fill = $context->Colors->Resolve($style->ForegroundColor, SKColors::Transparent());
            if ($fill->Alpha !== 0) {
                $effect->HasFill = true;
                $effect->Fill = $fill;
            }
        }

        if (!$effect->HasFontColor && $style->Font !== null && $style->Font->Color->A !== 0) {
            $effect->HasFontColor = true;
            $effect->FontColor = $context->Colors->Resolve($style->Font->Color, SKColors::Black());
        }

        if ($style->Font !== null && $style->Font->Bold) {
            $effect->HasFontBold = true;
        }

        if ($style->Font !== null && $style->Font->Italic) {
            $effect->HasFontItalic = true;
        }

        if ($style->Font !== null && $style->Font->Underline !== FontUnderlineType::None) {
            $effect->HasFontUnderline = true;
            $effect->FontUnderline = $style->Font->Underline;
        }
    }

    // --- Differential-format rule matchers -----------------------------------------------------

    private static function BuildMatcher($sheet, $context, $condition, $area)
    {
        switch ($condition->Type) {
            case FormatConditionType::ContainsText:
            {
                $needle = $condition->Formula1 ?? '';
                return function ($record) use ($needle) {
                    return stripos(self::TextOf($record), $needle) !== false;
                };
            }
            case FormatConditionType::NotContainsText:
            {
                $needle = $condition->Formula1 ?? '';
                return function ($record) use ($needle) {
                    return stripos(self::TextOf($record), $needle) === false;
                };
            }
            case FormatConditionType::BeginsWith:
            {
                $prefix = $condition->Formula1 ?? '';
                return function ($record) use ($prefix) {
                    return strncasecmp(self::TextOf($record), $prefix, strlen($prefix)) === 0;
                };
            }
            case FormatConditionType::EndsWith:
            {
                $suffix = $condition->Formula1 ?? '';
                return function ($record) use ($suffix) {
                    $text = self::TextOf($record);
                    if ($suffix === '') {
                        return true;
                    }
                    return strcasecmp(substr($text, -strlen($suffix)), $suffix) === 0;
                };
            }
            case FormatConditionType::TimePeriod:
                return self::BuildTimePeriodMatcher($condition->TimePeriod, $context);
            case FormatConditionType::CellValue:
                return self::BuildCellValueMatcher($condition);
            case FormatConditionType::DuplicateValues:
            case FormatConditionType::UniqueValues:
                return self::BuildUniquenessMatcher($sheet, $area, $condition->Type === FormatConditionType::DuplicateValues);
            case FormatConditionType::Top10:
            case FormatConditionType::Bottom10:
                return self::BuildTopMatcher($sheet, $area, $condition);
            case FormatConditionType::AboveAverage:
            case FormatConditionType::BelowAverage:
                return self::BuildAverageMatcher($sheet, $area, $condition);
            default:
                // Unmodelled rule types have no evaluator.
                return null;
        }
    }

    private static function BuildCellValueMatcher($condition)
    {
        $v1 = 0.0;
        $v2 = 0.0;
        $has1 = self::TryParseDouble($condition->Formula1, $v1);
        $has2 = self::TryParseDouble($condition->Formula2, $v2);
        $op = $condition->Operator;
        return function ($record) use ($has1, $v1, $has2, $v2, $op) {
            $x = 0.0;
            if (!self::TryNumber($record, $x)) {
                return false;
            }

            switch ($op) {
                case OperatorType::GreaterThan: return $has1 && $x > $v1;
                case OperatorType::GreaterOrEqual: return $has1 && $x >= $v1;
                case OperatorType::LessThan: return $has1 && $x < $v1;
                case OperatorType::LessOrEqual: return $has1 && $x <= $v1;
                case OperatorType::Equal: return $has1 && $x == $v1;
                case OperatorType::NotEqual: return $has1 && $x != $v1;
                case OperatorType::Between: return $has1 && $has2 && $x >= min($v1, $v2) && $x <= max($v1, $v2);
                case OperatorType::NotBetween: return $has1 && $has2 && ($x < min($v1, $v2) || $x > max($v1, $v2));
                default: return false;
            }
        };
    }

    private static function BuildTimePeriodMatcher($period, $context)
    {
        $today = (clone $context->ConditionalFormattingEvaluationDate)->setTime(0, 0, 0);
        $dateSystem = $context->Workbook->Settings->DateSystem;
        return function ($record) use ($today, $dateSystem, $period) {
            $serial = 0.0;
            if (!self::TryNumber($record, $serial)) {
                return false;
            }

            $date = self::SerialToDate($serial, $dateSystem);
            if ($date === null) {
                return false;
            }

            $d = (clone $date)->setTime(0, 0, 0);
            switch (strtolower($period ?? '')) {
                case 'today': return $d == $today;
                case 'yesterday': return $d == (clone $today)->modify('-1 day');
                case 'tomorrow': return $d == (clone $today)->modify('+1 day');
                case 'last7days': return $d <= $today && $d >= (clone $today)->modify('-6 days');
                case 'thismonth': return (int)$d->format('Y') === (int)$today->format('Y') && (int)$d->format('n') === (int)$today->format('n');
                case 'lastmonth':
                {
                    $m = (clone $today)->modify('first day of last month');
                    return (int)$d->format('Y') === (int)$m->format('Y') && (int)$d->format('n') === (int)$m->format('n');
                }
                case 'nextmonth':
                {
                    $m = (clone $today)->modify('first day of next month');
                    return (int)$d->format('Y') === (int)$m->format('Y') && (int)$d->format('n') === (int)$m->format('n');
                }
                default: return false;
            }
        };
    }

    private static function BuildUniquenessMatcher($sheet, $area, $wantDuplicates)
    {
        $counts = array();
        self::ForEachValue($sheet, $area, function ($record, $text) use (&$counts) {
            $key = strtolower($text);
            $n = array_key_exists($key, $counts) ? $counts[$key] : 0;
            $counts[$key] = $n + 1;
        });

        return function ($record) use ($counts, $wantDuplicates) {
            $text = self::ValueKey($record);
            if ($text === null) {
                return false;
            }

            $key = strtolower($text);
            $n = array_key_exists($key, $counts) ? $counts[$key] : 0;
            return $wantDuplicates ? $n > 1 : $n === 1;
        };
    }

    private static function BuildTopMatcher($sheet, $area, $condition)
    {
        $values = self::CollectNumbers($sheet, $area);
        if (count($values) === 0 || $condition->Rank <= 0) {
            return function ($record) {
                return false;
            };
        }

        $count = $condition->Percent
            ? (int)ceil(count($values) * ($condition->Rank / 100.0))
            : $condition->Rank;
        $count = max(1, min($count, count($values)));

        sort($values);
        $threshold = $condition->Top
            ? $values[count($values) - $count]   // keep the largest `count`
            : $values[$count - 1];             // keep the smallest `count`

        return function ($record) use ($condition, $threshold) {
            $x = 0.0;
            if (!self::TryNumber($record, $x)) {
                return false;
            }

            return $condition->Top ? $x >= $threshold : $x <= $threshold;
        };
    }

    private static function BuildAverageMatcher($sheet, $area, $condition)
    {
        $values = self::CollectNumbers($sheet, $area);
        if (count($values) === 0) {
            return function ($record) {
                return false;
            };
        }

        $sum = 0.0;
        foreach ($values as $v) $sum += $v;
        $mean = $sum / count($values);

        $cutoff = $mean;
        if ($condition->StandardDeviation > 0) {
            $sq = 0.0;
            foreach ($values as $v) $sq += ($v - $mean) * ($v - $mean);
            $sd = sqrt($sq / count($values));
            $cutoff = $condition->Above ? $mean + $condition->StandardDeviation * $sd : $mean - $condition->StandardDeviation * $sd;
        }

        return function ($record) use ($condition, $cutoff) {
            $x = 0.0;
            if (!self::TryNumber($record, $x)) {
                return false;
            }

            return $condition->Above ? $x > $cutoff : $x < $cutoff;
        };
    }

    // --- Color scale / data bar / icon set -----------------------------------------------------

    private function ApplyColorScale($sheet, $context, $condition, $area)
    {
        $values = self::CollectNumbers($sheet, $area);
        if (count($values) === 0) {
            return;
        }

        sort($values);
        $min = $values[0];
        $max = $values[count($values) - 1];
        $threeColor = $condition->ColorScaleCount >= 3 && $condition->MidColor->A !== 0;
        $mid = $threeColor ? self::Percentile($values, 50) : ($min + $max) / 2.0;

        $minColor = self::ToSk($context, $condition->MinColor, new SKColor(0xF8, 0x69, 0x6B));
        $midColor = self::ToSk($context, $condition->MidColor, new SKColor(0xFF, 0xEB, 0x84));
        $maxColor = self::ToSk($context, $condition->MaxColor, new SKColor(0x63, 0xBE, 0x7B));

        for ($r = $area->StartRow; $r <= $area->EndRow; $r++) {
            for ($c = $area->StartColumn; $c <= $area->EndColumn; $c++) {
                $key = (new CellAddress($r, $c))->ToString();
                $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
                $x = 0.0;
                if (!self::TryNumber($record, $x)) {
                    continue;
                }

                if (!$threeColor) {
                    $color = self::Lerp($minColor, $maxColor, self::Fraction($x, $min, $max));
                } elseif ($x <= $mid) {
                    $color = self::Lerp($minColor, $midColor, self::Fraction($x, $min, $mid));
                } else {
                    $color = self::Lerp($midColor, $maxColor, self::Fraction($x, $mid, $max));
                }

                $effect = $this->EffectFor($r, $c);
                $effect->HasFill = true;
                $effect->Fill = $color;
            }
        }
    }

    private function ApplyDataBar($sheet, $context, $condition, $area)
    {
        $values = self::CollectNumbers($sheet, $area);
        if (count($values) === 0) {
            return;
        }

        sort($values);
        $min = $values[0];
        $max = $values[count($values) - 1];
        if ($condition->DataBarAutoMinimum && $min >= 0.0) {
            $min = 0.0;
        }
        $color = self::ToSk($context, $condition->BarColor, new SKColor(0x63, 0x8E, 0xC6));

        for ($r = $area->StartRow; $r <= $area->EndRow; $r++) {
            for ($c = $area->StartColumn; $c <= $area->EndColumn; $c++) {
                $key = (new CellAddress($r, $c))->ToString();
                $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
                $x = 0.0;
                if (!self::TryNumber($record, $x)) {
                    continue;
                }

                $display = $condition->DataBarMinLength
                    + ($condition->DataBarMaxLength - $condition->DataBarMinLength) * self::Fraction($x, $min, $max);
                $effect = $this->EffectFor($r, $c);
                $bar = new DataBar();
                $bar->Fraction = $display;
                $bar->Color = $color;
                $bar->ShowBorder = $condition->ShowBorder;
                $bar->Gradient = $condition->DataBarGradient;
                $effect->Bar = $bar;
            }
        }
    }

    private function ApplyIconSet($sheet, $condition, $area)
    {
        $values = self::CollectNumbers($sheet, $area);
        if (count($values) === 0) {
            return;
        }

        sort($values);
        $min = $values[0];
        $max = $values[count($values) - 1];
        $iconCount = self::IconCount($condition->IconSetType);

        for ($r = $area->StartRow; $r <= $area->EndRow; $r++) {
            for ($c = $area->StartColumn; $c <= $area->EndColumn; $c++) {
                $key = (new CellAddress($r, $c))->ToString();
                $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
                $x = 0.0;
                if (!self::TryNumber($record, $x)) {
                    continue;
                }

                // OOXML's default percent thresholds are integer values (0, 33, 66 for three icons).
                $percent = self::Fraction($x, $min, $max) * 100.0;
                $bucket = 0;
                for ($index = 1; $index < $iconCount; $index++) {
                    if ($percent >= (100 * $index) / $iconCount) {
                        $bucket = $index;
                    }
                }

                $lookup = $condition->ReverseIcons ? $iconCount - 1 - $bucket : $bucket;
                $effect = $this->EffectFor($r, $c);
                $effect->IconGlyph = self::IconFor($condition->IconSetType, $lookup, $iconCount);
                if ($condition->ShowIconOnly) {
                    $effect->SuppressText = true;
                }
            }
        }
    }

    private static function IconFor($iconSetType, $index, $count)
    {
        self::InitializeColors();

        $type = strtolower($iconSetType ?? '');
        $isArrow = str_contains($type, 'arrow');
        $isTrafficLight = str_contains($type, 'trafficlight');

        // Colour ramp low->high: red, amber, green (amber fills the middle buckets).
        if ($index === 0) {
            $color = self::$ArrowRed;
        } elseif ($index === $count - 1) {
            $color = self::$ArrowGreen;
        } else {
            $color = self::$ArrowAmber;
        }

        $borderColor = SKColors::Transparent();
        if ($isTrafficLight) {
            if ($index === 0) {
                $color = self::$TrafficRed;
                $borderColor = self::$TrafficRedBorder;
            } elseif ($index === $count - 1) {
                $color = self::$TrafficGreen;
                $borderColor = self::$TrafficGreenBorder;
            } else {
                $color = self::$TrafficYellow;
                $borderColor = self::$TrafficYellowBorder;
            }
        }

        // Arrow direction sweeps from down (lowest) to up (highest) across the buckets.
        $angle = $count <= 1 ? 0.0 : -90.0 + 180.0 * $index / ($count - 1);
        $icon = new Icon();
        $icon->Color = $color;
        $icon->BorderColor = $borderColor;
        $icon->HasBorder = $isTrafficLight;
        $icon->AngleDeg = $angle;
        $icon->IsArrow = $isArrow;
        return $icon;
    }

    // --- Value / numeric helpers ---------------------------------------------------------------

    private static function ForEachValue($sheet, $area, $action)
    {
        for ($r = $area->StartRow; $r <= $area->EndRow; $r++) {
            for ($c = $area->StartColumn; $c <= $area->EndColumn; $c++) {
                $key = (new CellAddress($r, $c))->ToString();
                if (!array_key_exists($key, $sheet->Cells)) {
                    continue;
                }
                $record = $sheet->Cells[$key];

                $valueKey = self::ValueKey($record);
                if ($valueKey !== null) {
                    $action($record, $valueKey);
                }
            }
        }
    }

    private static function CollectNumbers($sheet, $area)
    {
        $list = array();
        for ($r = $area->StartRow; $r <= $area->EndRow; $r++) {
            for ($c = $area->StartColumn; $c <= $area->EndColumn; $c++) {
                $key = (new CellAddress($r, $c))->ToString();
                $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
                $x = 0.0;
                if (self::TryNumber($record, $x)) {
                    $list[] = $x;
                }
            }
        }

        return $list;
    }

    private static function ValueKey($record)
    {
        if ($record === null || $record->Value === null) {
            return null;
        }

        $x = 0.0;
        if (self::TryNumber($record, $x)) {
            return 'n:' . (string)$x;
        }

        $text = self::TextOf($record);
        return strlen($text) === 0 ? null : 's:' . $text;
    }

    private static function TextOf($record)
    {
        if ($record === null || $record->Value === null) {
            return '';
        }

        if (is_bool($record->Value)) {
            return $record->Value ? 'True' : 'False';
        }

        return (string)$record->Value;
    }

    public static function TryNumber($record, &$value)
    {
        $value = 0.0;
        if ($record === null || $record->Value === null) {
            return false;
        }

        if ($record->Kind === CellValueKind::String || $record->Kind === CellValueKind::Boolean || $record->Kind === CellValueKind::Error) {
            return false;
        }

        return self::TryToDouble($record->Value, $value);
    }

    private static function TryToDouble($value, &$result)
    {
        $result = 0.0;
        if (is_float($value)) { $result = (float)$value; return true; }
        if (is_int($value)) { $result = (float)$value; return true; }
        if ($value instanceof \DateTime) { $result = DateSerialConverter::ToSerial($value, DateSystem::Windows1900); return true; }
        return false;
    }

    private static function TryParseDouble($text, &$value)
    {
        $value = 0.0;
        if ($text === null || !is_numeric($text)) {
            return false;
        }

        $value = (float)$text;
        return true;
    }

    private static function SerialToDate($serial, $dateSystem)
    {
        try {
            return DateSerialConverter::FromSerial($serial, $dateSystem);
        } catch (\Throwable $exception) {
            return null;
        }
    }

    private static function Fraction($value, $min, $max)
    {
        if ($max <= $min) {
            return 0.0;
        }

        $f = ($value - $min) / ($max - $min);
        if ($f < 0.0) return 0.0;
        if ($f > 1.0) return 1.0;
        return $f;
    }

    private static function Percentile($sorted, $percent)
    {
        if (count($sorted) === 1) {
            return $sorted[0];
        }

        $rank = $percent / 100.0 * (count($sorted) - 1);
        $lo = (int)floor($rank);
        $hi = (int)ceil($rank);
        if ($lo === $hi) {
            return $sorted[$lo];
        }

        return $sorted[$lo] + ($rank - $lo) * ($sorted[$hi] - $sorted[$lo]);
    }

    private static function ToSk($context, $color, $fallback)
    {
        if ($color->A === 0 && $color->ThemeIndex === null) {
            return $fallback;
        }

        return $context->Colors->Resolve($color, $fallback);
    }

    private static function Lerp($a, $b, $t)
    {
        return new SKColor(
            $a->Red + (int)(($b->Red - $a->Red) * $t),
            $a->Green + (int)(($b->Green - $a->Green) * $t),
            $a->Blue + (int)(($b->Blue - $a->Blue) * $t));
    }

    private static function IconCount($iconSetType)
    {
        if ($iconSetType === null || $iconSetType === '') {
            return 3;
        }

        if ($iconSetType[0] === '4') return 4;
        if ($iconSetType[0] === '5') return 5;
        return 3;
    }

    private static function Key($row, $col)
    {
        return ($row << 20) | $col;
    }
}
