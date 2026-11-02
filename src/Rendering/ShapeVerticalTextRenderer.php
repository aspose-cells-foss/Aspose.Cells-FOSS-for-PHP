<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKFont;

final class ShapeVerticalTextRenderer
{
    private function __construct()
    {
    }

    public static function IsVerticalText($shape)
    {
        if ($shape === null || $shape->TextVerticalType === null || $shape->TextVerticalType === '') {
            return false;
        }

        return $shape->TextVerticalType === 'vert'
            || $shape->TextVerticalType === 'vert270'
            || $shape->TextVerticalType === 'wordArtVert'
            || $shape->TextVerticalType === 'wordArtVertRtl';
    }

    public static function Draw($canvas, $rect, $shape, $fontContext, $paint)
    {
        if ($shape->TextVerticalType === 'wordArtVert' || $shape->TextVerticalType === 'wordArtVertRtl') {
            self::DrawWordArtVerticalText($canvas, $rect, $shape, $fontContext, $paint);
            return;
        }

        $clockwise = $shape->TextVerticalType === 'vert270' || $shape->TextVerticalType === 'wordArtVertRtl';
        $canvas->Save();
        $canvas->Translate($rect->MidX, $rect->MidY);
        $canvas->RotateDegrees($clockwise ? 90.0 : -90.0);

        $metrics = $paint->FontMetrics;
        $width = $fontContext->Measure($shape->Text);
        $drawX = -$width / 2.0;
        $baseline = -($metrics->Ascent + $metrics->Descent) / 2.0;
        $fontObject = new SKFont($paint->Typeface, $paint->TextSize);
        try {
            $fontObject->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $shape->Text, $drawX, $baseline, $fontObject, $paint->Color);
        } finally {
            $fontObject->Dispose();
        }

        $canvas->Restore();
    }

    private static function DrawWordArtVerticalText($canvas, $rect, $shape, $fontContext, $paint)
    {
        if ($shape->Text === null || $shape->Text === '') {
            return;
        }

        $metrics = $paint->FontMetrics;
        $charAdvance = max((float)$shape->TextSizePt, $metrics->Descent - $metrics->Ascent);
        $availableHeight = max($charAdvance, $rect->Height - 2.0 * (float)$shape->TextInsetTopPt);
        $rows = max(1, (int)floor($availableHeight / $charAdvance));
        $columns = intdiv(strlen($shape->Text) + $rows - 1, $rows);
        $totalHeight = $charAdvance * min($rows, strlen($shape->Text));
        $currentY = $rect->MidY - $totalHeight * 0.5;

        $canvas->Save();
        $canvas->ClipRect($rect);
        $fontObject = new SKFont($paint->Typeface, $paint->TextSize);
        try {
            $fontObject->Subpixel = true;
            for ($i = 0; $i < strlen($shape->Text); $i++) {
                $ch = $shape->Text[$i];
                $column = intdiv($i, $rows);
                if ($shape->TextVerticalType === 'wordArtVertRtl') {
                    $column = $columns - 1 - $column;
                }
                $centerX = $rect->MidX + ($column - ($columns - 1) * 0.5) * $charAdvance;
                $width = $fontContext->Measure($ch);
                $drawX = $centerX - $width * 0.5;
                $baseline = $currentY + ($i % $rows) * $charAdvance - $metrics->Ascent;
                PdfTextPathRenderer::DrawText($canvas, $ch, $drawX, $baseline, $fontObject, $paint->Color);
            }
        } finally {
            $fontObject->Dispose();
        }
        $canvas->Restore();
    }
}
