<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Core\HorizontalAlignment;
use Aspose\Cells_FOSS\Core\VerticalAlignment;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;

final class RichTextCellRenderer
{
    private $context;

    public function __construct($context)
    {
        $this->context = $context;
    }

    public function CanRender($record, $style, $text)
    {
        if ($record === null || $record->RichTextRuns === null || count($record->RichTextRuns) === 0) {
            return false;
        }

        if ($style === null) {
            return false;
        }

        if ($style->Alignment->WrapText) {
            return false;
        }

        if ($style->Alignment->TextRotation !== 0) {
            return false;
        }

        if ($text === null || $text === '') {
            return false;
        }

        return strpos($text, "\r") === false && strpos($text, "\n") === false;
    }

    public function MeasureWidth($record, $style, $text, $useOverrideColor, $overrideColor)
    {
        $slices = $this->BuildSlices($record, $style, $text, $useOverrideColor, $overrideColor);
        $total = 0.0;
        for ($i = 0; $i < count($slices); $i++) {
            $total += $slices[$i]->WidthPt;
        }

        return $total;
    }

    public function TryDraw($canvas, $page, $rect, $clip, $record, $style, $horizontal, $useOverrideColor, $overrideColor)
    {
        $text = is_string($record->Value) ? $record->Value : null;
        if (!$this->CanRender($record, $style, $text)) {
            return false;
        }

        $slices = $this->BuildSlices($record, $style, $text, $useOverrideColor, $overrideColor);
        if (count($slices) === 0) {
            return false;
        }

        $totalWidth = 0.0;
        $maxAscent = 0.0;
        $maxDescent = 0.0;
        for ($i = 0; $i < count($slices); $i++) {
            $slice = $slices[$i];
            $totalWidth += $slice->WidthPt;
            if ($slice->AscentPt > $maxAscent) {
                $maxAscent = $slice->AscentPt;
            }

            if ($slice->DescentPt > $maxDescent) {
                $maxDescent = $slice->DescentPt;
            }
        }

        $lineHeight = $maxAscent + $maxDescent;
        if ($lineHeight <= 0.0) {
            $lineHeight = $this->context->GetFontContext($style->Font)->LineHeightPt;
        }

        switch ($style->Alignment->Vertical) {
            case VerticalAlignment::Top:
                $blockTop = $rect->Top + (float)SheetLayout::VerticalPaddingPt;
                break;
            case VerticalAlignment::Center:
                $blockTop = $rect->Top + ($rect->Height - $lineHeight) / 2.0;
                break;
            default:
                $blockTop = $rect->Bottom - (float)SheetLayout::VerticalPaddingPt - $lineHeight;
                break;
        }

        $padding = (float)SheetLayout::HorizontalPaddingPt;
        $indentOffset = self::IndentOffsetPt($this->context->GetFontContext($style->Font), $style->Alignment, $horizontal);
        switch ($horizontal) {
            case HorizontalAlignment::Center:
            case HorizontalAlignment::CenterContinuous:
                $x = $rect->Left + ($rect->Width - $totalWidth) / 2.0;
                break;
            case HorizontalAlignment::Right:
                $x = $rect->Right - $padding - $indentOffset - $totalWidth;
                break;
            default:
                $x = $rect->Left + $padding + $indentOffset;
                break;
        }

        $baseline = $blockTop + $maxAscent;
        $canvas->Save();
        $canvas->ClipRect($clip);

        $paint = new SKPaint();
        $paint->IsAntialias = true;

        for ($i = 0; $i < count($slices); $i++) {
            $slice = $slices[$i];
            $runs = $slice->FontContext->SplitRuns($slice->Text);
            $cursor = $x;
            $paint->Color = $slice->Color;

            for ($runIndex = 0; $runIndex < count($runs); $runIndex++) {
                $run = $runs[$runIndex];
                $font = new SKFont($run->Typeface, $slice->FontContext->SizePt);
                $font->Subpixel = true;
                PdfTextPathRenderer::DrawText($canvas, $run->Text, $cursor, $baseline, $font, $paint->Color);
                $font->Dispose();

                $this->DrawDecorations($canvas, $paint, $cursor, $baseline, $run->WidthPt, $slice->FontContext->SizePt, $slice->Font);
                $cursor += $run->WidthPt;
            }

            $x += $slice->WidthPt;
        }

        $paint->Dispose();
        $canvas->Restore();
        return true;
    }

    private static function IndentOffsetPt($fontContext, $alignment, $horizontal)
    {
        if ($fontContext === null || $alignment === null) {
            return 0.0;
        }

        if ($alignment->IndentLevel <= 0) {
            return 0.0;
        }

        if ($horizontal !== HorizontalAlignment::Left
            && $horizontal !== HorizontalAlignment::Right
            && $horizontal !== HorizontalAlignment::Distributed
            && $horizontal !== HorizontalAlignment::Justify) {
            return 0.0;
        }

        $digitWidth = $fontContext->Measure("0");
        $indentStep = max($fontContext->SizePt * 0.68, $digitWidth * 1.15);
        return $alignment->IndentLevel * $indentStep;
    }

    private function BuildSlices($record, $style, $text, $useOverrideColor, $overrideColor)
    {
        $slices = array();
        $currentIndex = 0;
        $runs = $record->RichTextRuns;
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            if ($run->StartIndex > $currentIndex) {
                $this->AddSlice($slices, substr($text, $currentIndex, $run->StartIndex - $currentIndex), $style->Font, $useOverrideColor, $overrideColor);
            }

            $runFont = $this->MergeFont($style->Font, $run->Font);
            $this->AddSlice($slices, substr($text, $run->StartIndex, $run->Length), $runFont, $useOverrideColor, $overrideColor);
            $currentIndex = $run->StartIndex + $run->Length;
        }

        if ($currentIndex < strlen($text)) {
            $this->AddSlice($slices, substr($text, $currentIndex), $style->Font, $useOverrideColor, $overrideColor);
        }

        return $slices;
    }

    private function AddSlice(&$slices, $text, $font, $useOverrideColor, $overrideColor)
    {
        if ($text === null || $text === '') {
            return;
        }

        $resolvedFont = $font !== null ? $font : new FontValue();
        $context = $this->context->GetFontContext($resolvedFont);
        $metrics = $this->MeasureMetrics($resolvedFont, $context);
        $slice = new RichTextSlice();
        $slice->Text = $text;
        $slice->Font = $resolvedFont;
        $slice->FontContext = $context;
        $slice->WidthPt = $context->Measure($text);
        $slice->AscentPt = $metrics[0];
        $slice->DescentPt = $metrics[1];
        $slice->Color = $useOverrideColor ? $overrideColor : $this->context->Colors->Resolve($resolvedFont->Color, SKColors::Black());
        $slices[] = $slice;
    }

    private function MeasureMetrics($font, $context)
    {
        $skFont = new SKFont($this->context->Fonts->Resolve($font), $context->SizePt);
        $metrics = $skFont->Metrics;
        $skFont->Dispose();
        return array(-$metrics->Ascent, $metrics->Descent);
    }

    private function MergeFont($baseFont, $richFont)
    {
        if ($baseFont === null && $richFont === null) {
            return new FontValue();
        }

        if ($baseFont === null) {
            return $richFont->Clone();
        }

        if ($richFont === null) {
            return $baseFont->Clone();
        }

        $merged = $baseFont->Clone();
        $merged->Name = $richFont->Name;
        $merged->Size = $richFont->Size;
        $merged->Bold = $richFont->Bold;
        $merged->Italic = $richFont->Italic;
        $merged->Underline = $richFont->Underline;
        $merged->StrikeThrough = $richFont->StrikeThrough;
        $merged->Color = $richFont->Color;
        $merged->Family = $richFont->Family;
        $merged->Scheme = $richFont->Scheme;
        return $merged;
    }

    private function DrawDecorations($canvas, $paint, $x, $baseline, $width, $fontSize, $font)
    {
        if ($font !== null && $font->Underline !== FontUnderlineType::None) {
            $linePaint = new SKPaint();
            $linePaint->Color = $paint->Color;
            $linePaint->Style = SKPaintStyle::Stroke;
            $linePaint->IsAntialias = true;
            $linePaint->StrokeWidth = max(0.6, $fontSize * 0.055);
            $underlineY = $baseline + $fontSize * 0.08;
            $canvas->DrawLine($x, $underlineY, $x + $width, $underlineY, $linePaint);
            $linePaint->Dispose();
        }

        if ($font !== null && $font->StrikeThrough) {
            $linePaint = new SKPaint();
            $linePaint->Color = $paint->Color;
            $linePaint->Style = SKPaintStyle::Stroke;
            $linePaint->IsAntialias = true;
            $linePaint->StrokeWidth = max(0.55, $fontSize * 0.05);
            $strikeY = $baseline - $fontSize * 0.28;
            $canvas->DrawLine($x, $strikeY, $x + $width, $strikeY, $linePaint);
            $linePaint->Dispose();
        }
    }
}
