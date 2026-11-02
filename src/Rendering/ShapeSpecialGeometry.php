<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;

final class ShapeSpecialGeometry
{
    private function __construct()
    {
    }

    public static function Build($shape, $rect)
    {
        if ($shape->CustomGeometry !== null) {
            return ShapeCustomGeometry::Build($shape->CustomGeometry, $rect);
        }

        if ($shape->Geometry === 'irregularSeal1') {
            // DrawingML irregularSeal1 vertices in its 21600 x 21600 coordinate space.
            $vertices = array(10800, 5800, 14522, 0, 14155, 5325, 18380, 4457,
                16702, 7315, 21097, 8137, 17607, 10475, 21600, 13290, 16837, 12942,
                18145, 18095, 14020, 14457, 13247, 19737, 10532, 14935, 8485, 21600,
                7715, 15627, 4762, 17617, 5667, 13937, 135, 14587, 3722, 11775,
                0, 8615, 4627, 7617, 370, 2295, 7312, 6320, 8352, 2295);
            $path = new SKPath();
            for ($i = 0; $i < count($vertices); $i += 2) {
                $x = $rect->Left + $rect->Width * $vertices[$i] / 21600.0;
                $y = $rect->Top + $rect->Height * $vertices[$i + 1] / 21600.0;
                if ($i === 0) {
                    $path->MoveTo($x, $y);
                } else {
                    $path->LineTo($x, $y);
                }
            }
            $path->Close();
            return $path;
        }

        if ($shape->Geometry === 'wedgeRectCallout') {
            return self::BuildCallout($shape->Adjusts, $rect);
        }

        if ($shape->Geometry === 'snip1Rect' || $shape->Geometry === 'snip2SameRect') {
            $top = min($rect->Width, $rect->Height) * max(0.0, min(0.5, ShapeCurveGeometry::Adjustment($shape, 0, 0.16667)));
            $both = $shape->Geometry === 'snip2SameRect';
            $bottom = $both ? min($rect->Width, $rect->Height) * max(0.0, min(0.5, ShapeCurveGeometry::Adjustment($shape, 1, 0.0))) : 0.0;
            $path = new SKPath();
            $path->MoveTo($rect->Left + ($both ? $top : 0.0), $rect->Top);
            $path->LineTo($rect->Right - $top, $rect->Top);
            $path->LineTo($rect->Right, $rect->Top + $top);
            $path->LineTo($rect->Right, $rect->Bottom - $bottom);
            $path->LineTo($rect->Right - $bottom, $rect->Bottom);
            $path->LineTo($rect->Left + $bottom, $rect->Bottom);
            $path->LineTo($rect->Left, $rect->Bottom - $bottom);
            $path->LineTo($rect->Left, $rect->Top + ($both ? $top : 0.0));
            $path->Close();
            return $path;
        }

        return ShapeExtendedGeometry::Build($shape, $rect);
    }

    private static function BuildCallout($adjusts, $r)
    {
        $dx = $adjusts !== null && count($adjusts) > 0 ? (float)$adjusts[0] : -0.20833;
        $dy = $adjusts !== null && count($adjusts) > 1 ? (float)$adjusts[1] : 0.625;
        $tip = new SKPoint($r->MidX + $r->Width * $dx, $r->MidY + $r->Height * $dy);
        $x1 = $r->Left + $r->Width * ($dx > 0.0 ? 7.0 : 2.0) / 12.0;
        $x2 = $r->Left + $r->Width * ($dx > 0.0 ? 10.0 : 5.0) / 12.0;
        $y1 = $r->Top + $r->Height * ($dy > 0.0 ? 7.0 : 2.0) / 12.0;
        $y2 = $r->Top + $r->Height * ($dy > 0.0 ? 10.0 : 5.0) / 12.0;
        $vertical = abs($dy) > abs($dx);
        $path = new SKPath();
        $path->MoveTo($r->Left, $r->Top);
        if ($vertical && $dy <= 0.0) {
            $path->LineTo($x1, $r->Top);
            $path->LineTo($tip);
            $path->LineTo($x2, $r->Top);
        }
        $path->LineTo($r->Right, $r->Top);
        if (!$vertical && $dx > 0.0) {
            $path->LineTo($r->Right, $y1);
            $path->LineTo($tip);
            $path->LineTo($r->Right, $y2);
        }
        $path->LineTo($r->Right, $r->Bottom);
        if ($vertical && $dy > 0.0) {
            $path->LineTo($x2, $r->Bottom);
            $path->LineTo($tip);
            $path->LineTo($x1, $r->Bottom);
        }
        $path->LineTo($r->Left, $r->Bottom);
        if (!$vertical && $dx <= 0.0) {
            $path->LineTo($r->Left, $y2);
            $path->LineTo($tip);
            $path->LineTo($r->Left, $y1);
        }
        $path->Close();
        return $path;
    }
}
