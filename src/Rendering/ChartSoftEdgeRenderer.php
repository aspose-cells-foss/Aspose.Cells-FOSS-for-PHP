<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKBlurStyle;
use Aspose\Cells_FOSS\Skia\SKMaskFilter;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKRect;

final class ChartSoftEdgeRenderer
{
    private function __construct()
    {
    }

    public static function DrawSolidFill($canvas, $area, $color, $radiusPt)
    {
        $radius = (float)min($radiusPt, min($area->Width, $area->Height) * 0.4);
        $inset = $radius * 0.8;
        $inner = new SKRect($area->Left + $inset, $area->Top + $inset,
            $area->Right - $inset, $area->Bottom - $inset);
        $blur = SKMaskFilter::CreateBlur(SKBlurStyle::Normal, $radius / 3.0);
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Fill;
        $paint->Color = $color;
        $paint->IsAntialias = true;
        $paint->MaskFilter = $blur;
        $canvas->Save();
        $canvas->ClipRect($area);
        $canvas->DrawRect($inner, $paint);
        $canvas->Restore();
        $paint->Dispose();
        $blur->Dispose();
    }
}
