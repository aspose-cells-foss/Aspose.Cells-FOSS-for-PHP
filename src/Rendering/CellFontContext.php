<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKRect;

final class CellFontContext
{
    private $registry;

    private $primary;

    private $sizePt;

    private $fonts = array();

    public function __construct($registry, $font)
    {
        $this->registry = $registry;
        $this->primary = $registry->Resolve($font);
        $this->sizePt = (float)($font !== null ? $font->Size : 11.0);
        if ($this->sizePt <= 0.0) {
            $this->sizePt = 11.0;
        }
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'SizePt') {
            return $this->sizePt;
        }

        if ($propertyName === 'LineHeightPt') {
            return $this->sizePt * 1.2;
        }

        if ($propertyName === 'WrappedLineHeightPt') {
            $metrics = $this->FontFor($this->primary)->Metrics;
            return max($this->sizePt * 1.2, -$metrics->Ascent + $metrics->Descent + $metrics->Leading + 1.0);
        }

        if ($propertyName === 'TextHeightPt') {
            $metrics = $this->FontFor($this->primary)->Metrics;
            return -$metrics->Ascent + $metrics->Descent;
        }

        return null;
    }

    public function MeasureInkBounds($text, &$width, &$height)
    {
        $width = 0.0;
        $height = 0.0;
        if ($text === null || $text === '') {
            return;
        }

        $bounds = new SKRect();
        $paint = new SKPaint($this->FontFor($this->primary));
        try {
            $paint->MeasureText($text, $bounds);
        } finally {
            $paint->Dispose();
        }

        $width = $bounds->Width;
        $height = $bounds->Height;
        if ($height <= 0.0) {
            $height = $this->TextHeightPt;
        }
    }

    private function FontFor($typeface)
    {
        for ($i = 0; $i < count($this->fonts); $i++) {
            if ($this->fonts[$i][0] === $typeface) {
                return $this->fonts[$i][1];
            }
        }

        $font = new SKFont($typeface, $this->sizePt);
        $font->Subpixel = true;
        $this->fonts[] = array($typeface, $font);
        return $font;
    }

    public function SplitRuns($text)
    {
        $runs = array();
        if ($text === null || $text === '') {
            return $runs;
        }

        $builder = '';
        $currentFace = null;

        $index = 0;
        $length = strlen($text);
        while ($index < $length) {
            $consumed = self::ReadCodepoint($text, $index, $codepoint);
            $face = $this->registry->ResolveForCodepoint($this->primary, $codepoint);

            if ($currentFace === null) {
                $currentFace = $face;
            } elseif ($face !== $currentFace) {
                $runs[] = $this->MakeRun($currentFace, $builder);
                $builder = '';
                $currentFace = $face;
            }

            $builder .= substr($text, $index, $consumed);
            $index += $consumed;
        }

        if ($builder !== '') {
            $runs[] = $this->MakeRun($currentFace, $builder);
        }

        return $runs;
    }

    private function MakeRun($typeface, $text)
    {
        $font = $this->FontFor($typeface);
        $paint = new SKPaint($font);
        try {
            $run = new GlyphRun();
            $run->Typeface = $typeface;
            $run->Text = $text;
            $run->WidthPt = $paint->MeasureText($text);
            return $run;
        } finally {
            $paint->Dispose();
        }
    }

    public function Measure($text)
    {
        $runs = $this->SplitRuns($text);
        $total = 0.0;
        for ($i = 0; $i < count($runs); $i++) {
            $total += $runs[$i]->WidthPt;
        }

        return $total;
    }

    public function SplitPdfRuns($text)
    {
        $runs = $this->SplitRuns($text);
        for ($i = 0; $i < count($runs); $i++) {
            $font = new SKFont($runs[$i]->Typeface, 1000.0);
            try {
                $font->Subpixel = true;
                $paint = new SKPaint($font);
                try {
                    $width = 0.0;
                    $index = 0;
                    $length = strlen($runs[$i]->Text);
                    while ($index < $length) {
                        $count = self::ReadCodepoint($runs[$i]->Text, $index, $codepoint);
                        $width += $paint->MeasureText(substr($runs[$i]->Text, $index, $count));
                        $index += $count;
                    }

                    $runs[$i]->WidthPt = $width * $this->sizePt / 1000.0;
                } finally {
                    $paint->Dispose();
                }
            } finally {
                $font->Dispose();
            }
        }

        return $runs;
    }

    public function MeasurePdfText($text)
    {
        $runs = $this->SplitPdfRuns($text);
        $width = 0.0;
        for ($i = 0; $i < count($runs); $i++) {
            $width += $runs[$i]->WidthPt;
        }

        return $width;
    }

    public function WrapLines($text, $maxWidthPt, $pdfMetrics = false)
    {
        $lines = array();
        if ($text === null || $text === '') {
            $lines[] = '';
            return $lines;
        }

        $paragraphs = explode("\n", str_replace("\r\n", "\n", str_replace("\r", "\n", $text)));
        for ($p = 0; $p < count($paragraphs); $p++) {
            $this->WrapParagraph($paragraphs[$p], $maxWidthPt, $lines, $pdfMetrics);
        }

        return $lines;
    }

    private function MeasureForWrap($text, $pdfMetrics)
    {
        return $pdfMetrics ? $this->MeasurePdfText($text) : $this->Measure($text);
    }

    private function WrapParagraph($paragraph, $maxWidthPt, &$lines, $pdfMetrics)
    {
        if ($maxWidthPt <= 0.0 || strlen($paragraph) === 0) {
            $lines[] = $paragraph;
            return;
        }

        $current = '';
        $tokens = self::Tokenize($paragraph);
        for ($t = 0; $t < count($tokens); $t++) {
            $token = $tokens[$t];
            $candidate = $current === '' ? $token : $current . $token;
            if ($this->MeasureForWrap(rtrim($candidate), $pdfMetrics) <= $maxWidthPt) {
                $current .= $token;
                continue;
            }

            if ($current !== '') {
                $lines[] = rtrim($current);
                $current = '';
            }

            if ($this->MeasureForWrap(rtrim($token), $pdfMetrics) > $maxWidthPt) {
                $this->HardBreak($token, $maxWidthPt, $lines, $current, $pdfMetrics);
            } else {
                $current .= $token;
            }
        }

        $lines[] = rtrim($current);
    }

    private function HardBreak($token, $maxWidthPt, &$lines, &$current, $pdfMetrics)
    {
        $length = strlen($token);
        for ($i = 0; $i < $length; $i++) {
            $ch = $token[$i];
            $candidate = $current . $ch;
            if (ctype_space($ch)) {
                $current .= $ch;
                continue;
            }
            if ($current !== '' && $this->MeasureForWrap($candidate, $pdfMetrics) > $maxWidthPt) {
                $lines[] = $current;
                $current = '';
            }

            $current .= $ch;
        }
    }

    private static function Tokenize($paragraph)
    {
        $tokens = array();
        $builder = '';
        $length = strlen($paragraph);
        for ($i = 0; $i < $length; $i++) {
            $ch = $paragraph[$i];
            $builder .= $ch;
            if ($ch === ' ' || $ch === "\t") {
                $tokens[] = $builder;
                $builder = '';
            }
        }

        if ($builder !== '') {
            $tokens[] = $builder;
        }

        return $tokens;
    }

    private static function ReadCodepoint($text, $index, &$codepoint)
    {
        $first = ord($text[$index]);
        $length = strlen($text);

        if ($first < 0x80) {
            $codepoint = $first;
            return 1;
        }

        if ($first >= 0xC0 && $first < 0xE0 && $index + 1 < $length) {
            $codepoint = (($first & 0x1F) << 6) | (ord($text[$index + 1]) & 0x3F);
            return 2;
        }

        if ($first >= 0xE0 && $first < 0xF0 && $index + 2 < $length) {
            $codepoint = (($first & 0x0F) << 12) | ((ord($text[$index + 1]) & 0x3F) << 6) | (ord($text[$index + 2]) & 0x3F);
            return 3;
        }

        if ($first >= 0xF0 && $index + 3 < $length) {
            $codepoint = (($first & 0x07) << 18) | ((ord($text[$index + 1]) & 0x3F) << 12)
                | ((ord($text[$index + 2]) & 0x3F) << 6) | (ord($text[$index + 3]) & 0x3F);
            return 4;
        }

        $codepoint = $first;
        return 1;
    }

    public function Dispose()
    {
        for ($i = 0; $i < count($this->fonts); $i++) {
            $this->fonts[$i][1]->Dispose();
        }

        $this->fonts = array();
    }
}
