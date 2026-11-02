<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKRect;

final class ShapeExtendedGeometry
{
    private function __construct()
    {
    }

    public static function Build($shape, $r)
    {
        if ($r->Width <= 0.0 || $r->Height <= 0.0) {
            return null;
        }
        if ($shape->Geometry === 'leftArrowCallout') {
            return self::LeftArrowCallout($shape, $r);
        }
        if ($shape->Geometry === 'mathNotEqual') {
            return self::NotEqual($shape, $r);
        }
        if ($shape->Geometry === 'flowChartMagneticDrum') {
            $path = new SKPath();
            $path->MoveTo($r->Left + $r->Width / 6.0, $r->Top);
            $path->LineTo($r->Right - $r->Width / 6.0, $r->Top);
            $path->ArcTo(new SKRect($r->Right - $r->Width / 3.0, $r->Top, $r->Right, $r->Bottom), -90.0, 180.0, false);
            $path->LineTo($r->Left + $r->Width / 6.0, $r->Bottom);
            $path->ArcTo(new SKRect($r->Left, $r->Top, $r->Left + $r->Width / 3.0, $r->Bottom), 90.0, 180.0, false);
            $path->Close();
            return $path;
        }
        return null;
    }

    public static function Adjustment($shape, $index, $fallback)
    {
        $value = ShapeCurveGeometry::Adjustment($shape, $index, $fallback);
        return is_nan($value) || is_infinite($value) ? $fallback : $value;
    }

    private static function Clamp($value, $maximum)
    {
        return max(0.0, min($maximum, $value));
    }

    private static function LeftArrowCallout($shape, $r)
    {
        $ss = min($r->Width, $r->Height);
        $a2 = self::Clamp(self::Adjustment($shape, 1, 0.25), $r->Height / (2.0 * $ss));
        $a1 = self::Clamp(self::Adjustment($shape, 0, 0.25), $a2 * 2.0);
        $a3 = self::Clamp(self::Adjustment($shape, 2, 0.25), $r->Width / $ss);
        $a4 = self::Clamp(self::Adjustment($shape, 3, 0.64977), 1.0 - $a3 * $ss / $r->Width);
        $x1 = $r->Left + $ss * $a3;
        $x2 = $r->Right - $r->Width * $a4;
        $y1 = $r->MidY - $ss * $a2;
        $y2 = $r->MidY - $ss * $a1 / 2.0;
        $y3 = $r->MidY + $ss * $a1 / 2.0;
        $y4 = $r->MidY + $ss * $a2;
        return self::Polygon(array($r->Left, $r->MidY, $x1, $y1, $x1, $y2, $x2, $y2, $x2, $r->Top,
            $r->Right, $r->Top, $r->Right, $r->Bottom, $x2, $r->Bottom, $x2, $y3, $x1, $y3, $x1, $y4));
    }

    private static function NotEqual($shape, $r)
    {
        $w = $r->Width;
        $h = $r->Height;
        $a1 = self::Clamp(self::Adjustment($shape, 0, 0.2352), 0.5);
        $a3 = self::Clamp(self::Adjustment($shape, 2, 0.1176), 1.0 - 2.0 * $a1);
        // The second adjustment is an angle in 60000ths of a degree;
        // parsed adjustments retain the common /100000 scale.
        $angle = max(42.0, min(66.0, self::Adjustment($shape, 1, 66.0))) / 0.6 - 90.0;
        $dx = $h / 2.0 * (float)tan($angle * M_PI / 180.0);
        $length = (float)sqrt($dx * $dx + $h * $h / 4.0);
        $thickness = $h * $a1;
        $bandWidth = $length * $thickness / ($h / 2.0);
        $x1 = $w * (0.5 - 0.36745);
        $x8 = $w - $x1;
        $y2 = $h * (0.5 - $a3 / 2.0);
        $y3 = $h - $y2;
        $y1 = $y2 - $thickness;
        $y4 = $y3 + $thickness;
        $x7 = $w / 2.0 + $dx - $bandWidth / 2.0;
        $x6 = $x7 - $dx * $y1 / ($h / 2.0);
        $x5 = $x7 - $dx * $y2 / ($h / 2.0);
        $x4 = $x7 - $dx * $y3 / ($h / 2.0);
        $x3 = $x7 - $dx * $y4 / ($h / 2.0);
        $dx7 = $thickness * $h / 2.0 / $length;
        $dy = $thickness * $dx / $length;
        $rx = $angle > 0.0 ? $x7 + $dx7 : $x7 + $bandWidth;
        $lx = $angle > 0.0 ? $x7 : $x7 + $bandWidth - $dx7;
        $ry = $angle > 0.0 ? $dy : 0.0;
        $ly = $angle > 0.0 ? 0.0 : -$dy;
        $points = array($x1, $y1, $x6, $y1, $lx, $ly, $rx, $ry, $x6 + $bandWidth, $y1,
            $x8, $y1, $x8, $y2, $x5 + $bandWidth, $y2, $x4 + $bandWidth, $y3, $x8, $y3, $x8, $y4,
            $x3 + $bandWidth, $y4, $w - $lx, $h - $ly, $w - $rx, $h - $ry, $x3, $y4, $x1, $y4, $x1, $y3, $x4, $y3, $x5, $y2, $x1, $y2);
        for ($i = 0; $i < count($points); $i += 2) {
            $points[$i] += $r->Left;
            $points[$i + 1] += $r->Top;
        }
        return self::Polygon($points);
    }

    public static function Polygon($points)
    {
        $path = new SKPath();
        $path->MoveTo($points[0], $points[1]);
        for ($i = 2; $i < count($points); $i += 2) {
            $path->LineTo($points[$i], $points[$i + 1]);
        }
        $path->Close();
        return $path;
    }
}
