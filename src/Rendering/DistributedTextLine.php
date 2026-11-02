<?php

namespace Aspose\Cells_FOSS\Rendering;

final class DistributedTextLine
{
    public $Words;
    public $Offsets;

    public static function UnwrappedWidth($font, $innerWidth, $alignment)
    {
        $indentStep = max($font->SizePt * 0.68, $font->Measure('0') * 1.15);
        $indent = max(0, $alignment->IndentLevel + $alignment->RelativeIndent) * $indentStep;
        return max(1.0, $innerWidth - 2.0 * $indent);
    }

    public static function Build($font, $line, $width, $trailingSeparator)
    {
        $words = preg_split('/[ \t]+/', $line, -1, PREG_SPLIT_NO_EMPTY);
        if ($words === false) {
            $words = array();
        }
        $result = new DistributedTextLine();
        $result->Words = $words;
        $result->Offsets = array_fill(0, count($words), 0.0);
        if (count($words) === 0) {
            return $result;
        }
        $widths = array_fill(0, count($words), 0.0);
        $total = 0.0;
        for ($i = 0; $i < count($words); $i++) {
            $widths[$i] = $font->MeasurePdfText($words[$i]);
            $total += $widths[$i];
        }
        if (count($words) === 1) {
            $result->Offsets[0] = ($width - $total) / 2.0;
            return $result;
        }
        $gaps = count($words) - 1 + ($trailingSeparator ? 1 : 0);
        $gap = max(0.0, $width - $total) / $gaps;
        for ($i = 1; $i < count($words); $i++) {
            $result->Offsets[$i] = $result->Offsets[$i - 1] + $widths[$i - 1] + $gap;
        }
        return $result;
    }
}
