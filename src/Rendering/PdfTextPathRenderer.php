<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKStrokeJoin;

final class PdfTextPathRenderer
{
    private function __construct()
    {
    }

    public static function DrawText($canvas, $text, $x, $baseline, $font, $color, $style = SKPaintStyle::Fill, $strokeWidth = 0.0)
    {
        if ($canvas === null || $font === null || $text === null || $text === '') {
            return;
        }

        $paint = new SKPaint($font);
        try {
            $paint->Color = $color;
            $paint->Style = $style;
            $paint->StrokeWidth = $strokeWidth;
            $paint->StrokeJoin = SKStrokeJoin::Round;
            $paint->IsAntialias = true;

            $path = $paint->GetTextPath($text, $x, $baseline);
            if ($path !== null) {
                $canvas->DrawPath($path, $paint);
                $path->Dispose();
            }
        } finally {
            $paint->Dispose();
        }
    }
}
