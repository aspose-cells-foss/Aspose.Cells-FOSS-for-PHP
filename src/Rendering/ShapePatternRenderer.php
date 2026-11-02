<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKClipOperation;
use Aspose\Cells_FOSS\Skia\SKPaint;

final class ShapePatternRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $rect, $outline, $shape)
    {
        if ($shape->FillPattern !== 'pct5') {
            return;
        }
        // DrawingML pct5: two 2x2 dots in a 16x16 tile, rendered at 192 dpi.
        $tile = 6.0;
        $dot = 0.75;
        $canvas->Save();
        $canvas->ClipPath($outline, SKClipOperation::Intersect, true);
        $paint = new SKPaint();
        $paint->Color = $shape->PatternForeground;
        $paint->IsAntialias = true;
        for ($y = $rect->Top; $y < $rect->Bottom; $y += $tile) {
            for ($x = $rect->Left; $x < $rect->Right; $x += $tile) {
                $canvas->DrawRect($x, $y, $dot, $dot, $paint);
                $canvas->DrawRect($x + $tile / 2.0, $y + $tile / 2.0, $dot, $dot, $paint);
            }
        }
        $paint->Dispose();
        $canvas->Restore();
    }
}
