<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKStrokeJoin;

final class ShapePresetDetails
{
    private function __construct()
    {
    }

    public static function DrawFill($canvas, $r, $shape)
    {
        if ($shape->Geometry !== 'bevel') {
            return false;
        }
        $inset = self::BevelInset($shape, $r);
        $x1 = $r->Left + $inset;
        $x2 = $r->Right - $inset;
        $y1 = $r->Top + $inset;
        $y2 = $r->Bottom - $inset;
        // Disjoint faces: do not paint a base rectangle under translucent facets.
        self::Face($canvas, $shape->FillColor, array($x1, $y1, $x2, $y1, $x2, $y2, $x1, $y2));
        self::Face($canvas, self::FaceColor($shape->FillColor, 0.2), array($r->Left, $r->Top, $r->Right, $r->Top, $x2, $y1, $x1, $y1));
        self::Face($canvas, self::FaceColor($shape->FillColor, -0.2), array($r->Left, $r->Bottom, $x1, $y2, $x2, $y2, $r->Right, $r->Bottom));
        self::Face($canvas, self::FaceColor($shape->FillColor, 0.4), array($r->Left, $r->Top, $x1, $y1, $x1, $y2, $r->Left, $r->Bottom));
        self::Face($canvas, self::FaceColor($shape->FillColor, -0.4), array($r->Right, $r->Top, $r->Right, $r->Bottom, $x2, $y2, $x2, $y1));
        return true;
    }

    public static function DrawLines($canvas, $r, $shape)
    {
        if (!$shape->HasLine) {
            return;
        }

        $path = new SKPath();
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $shape->LineColor;
        $paint->StrokeWidth = max(0.5, $shape->LineWidthPt);
        $paint->IsAntialias = true;
        $paint->StrokeJoin = SKStrokeJoin::Miter;

        if ($shape->Geometry === 'flowChartMagneticDrum') {
            $path->ArcTo(new SKRect($r->Right - $r->Width / 3.0, $r->Top, $r->Right, $r->Bottom), 90.0, 180.0, true);
        } elseif ($shape->Geometry === 'accentBorderCallout3') {
            $defaults = array(0.1875, -0.08333, 0.1875, -0.16667, 1.0, -0.16667, 1.12963, -0.08333);
            $points = array();
            for ($i = 0; $i < 4; $i++) {
                $points[$i] = new SKPoint($r->Left + $r->Width * ShapeExtendedGeometry::Adjustment($shape, $i * 2 + 1, $defaults[$i * 2 + 1]),
                    $r->Top + $r->Height * ShapeExtendedGeometry::Adjustment($shape, $i * 2, $defaults[$i * 2]));
            }
            $path->MoveTo($points[0]->X, $r->Top);
            $path->LineTo($points[0]->X, $r->Bottom);
            $path->MoveTo($points[0]);
            for ($i = 1; $i < count($points); $i++) {
                $path->LineTo($points[$i]);
            }
        } elseif ($shape->Geometry === 'bevel') {
            $inset = self::BevelInset($shape, $r);
            $inner = new SKRect($r->Left + $inset, $r->Top + $inset, $r->Right - $inset, $r->Bottom - $inset);
            $path->AddRect($inner);
            $path->MoveTo($r->Left, $r->Top);
            $path->LineTo($inner->Left, $inner->Top);
            $path->MoveTo($r->Right, $r->Top);
            $path->LineTo($inner->Right, $inner->Top);
            $path->MoveTo($r->Right, $r->Bottom);
            $path->LineTo($inner->Right, $inner->Bottom);
            $path->MoveTo($r->Left, $r->Bottom);
            $path->LineTo($inner->Left, $inner->Bottom);
        } else {
            $path->Dispose();
            $paint->Dispose();
            return;
        }
        $canvas->DrawPath($path, $paint);
        $path->Dispose();
        $paint->Dispose();
    }

    private static function BevelInset($shape, $r)
    {
        return min($r->Width, $r->Height) * max(0.0, min(0.5, ShapeExtendedGeometry::Adjustment($shape, 0, 0.125)));
    }

    private static function Face($canvas, $color, $points)
    {
        $path = ShapeExtendedGeometry::Polygon($points);
        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->IsAntialias = true;
        $canvas->DrawPath($path, $paint);
        $path->Dispose();
        $paint->Dispose();
    }

    private static function FaceColor($color, $amount)
    {
        return new SKColor(self::Channel($color->Red, $amount), self::Channel($color->Green, $amount), self::Channel($color->Blue, $amount), $color->Alpha);
    }

    private static function Channel($value, $amount)
    {
        return (int)round($amount >= 0.0 ? $value + (255 - $value) * $amount : $value * (1.0 + $amount), 0, PHP_ROUND_HALF_EVEN);
    }
}
