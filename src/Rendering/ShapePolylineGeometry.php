<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPoint;

final class ShapePolylineGeometry
{
    private function __construct()
    {
    }

    public static function Build($start, $end, $shape)
    {
        $a1 = ShapeCurveGeometry::Adjustment($shape, 0, 0.5);
        $a2 = ShapeCurveGeometry::Adjustment($shape, 1, 0.5);
        switch ($shape->Geometry) {
            case 'bentConnector2':
                $points = array(new SKPoint(0.0, 0.0), new SKPoint(1.0, 0.0), new SKPoint(1.0, 1.0));
                break;
            case 'bentConnector3':
                $points = array(new SKPoint(0.0, 0.0), new SKPoint($a1, 0.0), new SKPoint($a1, 1.0), new SKPoint(1.0, 1.0));
                break;
            case 'bentConnector4':
                $points = array(new SKPoint(0.0, 0.0), new SKPoint($a1, 0.0), new SKPoint($a1, $a2), new SKPoint(1.0, $a2), new SKPoint(1.0, 1.0));
                break;
            default:
                return array($start, $end);
        }
        return ShapeCurveGeometry::Map($points, $start, $end, $shape);
    }
}
