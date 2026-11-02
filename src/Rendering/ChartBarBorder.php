<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKStrokeJoin;

final class ChartBarBorder
{
    const C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}';

    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function Apply($element, $series)
    {
        $spPr = $element->Element(self::C . 'spPr');
        $line = $spPr !== null ? $spPr->Element(self::A . 'ln') : null;
        $series->HasBarBorder = $line !== null && $series->HasVisibleLine;
        if ($line !== null && $line->Attribute('w') === null) {
            $series->LineWidthPt = 0.75;
        }
    }

    public static function Draw($canvas, $bounds, $series)
    {
        if (!$series->HasBarBorder || !$series->HasVisibleLine || $bounds->IsEmpty) {
            return;
        }

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $series->LineColor;
        $paint->StrokeWidth = (float)$series->LineWidthPt;
        $paint->IsAntialias = true;
        $paint->StrokeJoin = SKStrokeJoin::Miter;
        $canvas->DrawRect($bounds, $paint);
        $paint->Dispose();
    }
}
