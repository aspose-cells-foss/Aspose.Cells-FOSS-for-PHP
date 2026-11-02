<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPoint;

final class ShapeCurveGeometry
{
    private function __construct()
    {
    }

    public static function Build($start, $end, $shape)
    {
        if ($shape->Geometry === 'curvedConnector4') {
            $x2 = self::Adjustment($shape, 0, 0.5);
            $y4 = self::Adjustment($shape, 1, 0.5);
            $x3 = (1.0 + $x2) / 2.0;
            return self::Map(array(new SKPoint(0.0, 0.0), new SKPoint($x2 / 2.0, 0.0),
                new SKPoint($x2, $y4 / 4.0), new SKPoint($x2, $y4 / 2.0),
                new SKPoint($x2, $y4 * 0.75), new SKPoint(($x2 + $x3) / 2.0, $y4), new SKPoint($x3, $y4),
                new SKPoint(($x3 + 1.0) / 2.0, $y4), new SKPoint(1.0, (1.0 + $y4) / 2.0), new SKPoint(1.0, 1.0)), $start, $end, $shape);
        }
        if ($shape->Geometry !== 'curvedConnector3') {
            return array($start, new SKPoint(($start->X + $end->X) / 2.0, $start->Y),
                new SKPoint($end->X, ($start->Y + $end->Y) / 2.0), $end);
        }

        // The three-leg preset consists of two cubics, with horizontal tangents
        // at both ends and a vertical tangent at the adjustable middle joint.
        $adj = $shape->Adjusts !== null && count($shape->Adjusts) > 0 ? (float)$shape->Adjusts[0] : 0.5;
        $normalized = array(new SKPoint(0.0, 0.0), new SKPoint($adj / 2.0, 0.0),
            new SKPoint($adj, 0.25), new SKPoint($adj, 0.5), new SKPoint($adj, 0.75),
            new SKPoint((1.0 + $adj) / 2.0, 1.0), new SKPoint(1.0, 1.0));
        return self::Map($normalized, $start, $end, $shape);
    }

    public static function Adjustment($shape, $index, $fallback)
    {
        return $shape->Adjusts !== null && count($shape->Adjusts) > $index ? (float)$shape->Adjusts[$index] : $fallback;
    }

    public static function Map($normalized, $start, $end, $shape)
    {
        $vertical = abs(sin($shape->RotationDeg * M_PI / 180.0)) > 0.707;
        for ($i = 0; $i < count($normalized); $i++) {
            $x = $vertical ? $normalized[$i]->Y : $normalized[$i]->X;
            $y = $vertical ? $normalized[$i]->X : $normalized[$i]->Y;
            $normalized[$i] = new SKPoint($start->X + ($end->X - $start->X) * $x, $start->Y + ($end->Y - $start->Y) * $y);
        }

        return $normalized;
    }
}
