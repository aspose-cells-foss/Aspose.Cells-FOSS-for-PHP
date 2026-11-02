<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKFont;

final class ShapeWrappedTextRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $rect, $shape, $fontContext, $paint)
    {
        $availableWidth = max(1.0, $rect->Width - (float)$shape->TextInsetLeftPt * 2.0);
        $lines = array();
        $words = array();
        foreach (explode(' ', $shape->Text) as $word) {
            if ($word !== '') {
                $words[] = $word;
            }
        }
        $current = '';
        foreach ($words as $word) {
            $candidate = strlen($current) === 0 ? $word : $current . ' ' . $word;
            if (strlen($current) > 0 && $fontContext->Measure($candidate) > $availableWidth) {
                $lines[] = $current;
                $current = $word;
            } else {
                $current = $candidate;
            }
        }

        if (strlen($current) > 0) {
            $lines[] = $current;
        }

        if (count($lines) === 0) {
            return;
        }

        $metrics = $paint->FontMetrics;
        $lineHeight = $paint->TextSize * 1.2;
        $textHeight = (count($lines) - 1) * $lineHeight + $metrics->Descent - $metrics->Ascent;
        if ($shape->TextVAnchor === 't') {
            $baseline = $rect->Top + (float)$shape->TextInsetTopPt - $metrics->Ascent;
        } elseif ($shape->TextVAnchor === 'b') {
            $baseline = $rect->Bottom - (float)$shape->TextInsetTopPt - $textHeight - $metrics->Ascent;
        } else {
            $baseline = $rect->MidY - $textHeight / 2.0 - $metrics->Ascent;
        }

        $canvas->Save();
        $canvas->ClipRect($rect);
        $fontObject = new SKFont($paint->Typeface, $paint->TextSize);
        try {
            $fontObject->Subpixel = true;
            foreach ($lines as $line) {
                $lineWidth = $fontContext->Measure($line);
                if ($shape->TextHAlign === 'l') {
                    $x = $rect->Left + (float)$shape->TextInsetLeftPt;
                } elseif ($shape->TextHAlign === 'r') {
                    $x = $rect->Right - (float)$shape->TextInsetLeftPt - $lineWidth;
                } else {
                    $x = $rect->MidX - $lineWidth / 2.0;
                }
                PdfTextPathRenderer::DrawText($canvas, $line, $x, $baseline, $fontObject, $paint->Color);
                $baseline += $lineHeight;
            }
        } finally {
            $fontObject->Dispose();
        }
        $canvas->Restore();
    }
}
