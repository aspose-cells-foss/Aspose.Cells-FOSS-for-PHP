<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKPoint;

final class Column3DWallRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $area, $chart, $wallDepth)
    {
        if ($canvas === null || $chart === null
            || ($chart->SideWallImageData === null && $chart->BackWallImageData === null)) {
            return;
        }

        $backBottomLeft = Column3DGeometry::GridPoint($area, 0.0, 0.0);
        $backBottomRight = Column3DGeometry::GridPoint($area, 1.0, 0.0);
        $backTopLeft = Column3DGeometry::GridPoint($area, 0.0, 1.0);
        $backTopRight = Column3DGeometry::GridPoint($area, 1.0, 1.0);
        $frontBottomLeft = self::Offset($backBottomLeft, $wallDepth);
        $frontTopLeft = self::Offset($backTopLeft, $wallDepth);
        Area3DWallFillRenderer::Draw($canvas, $frontBottomLeft, $frontTopLeft,
            $backBottomLeft, $backTopLeft, $backBottomRight, $backTopRight,
            $chart->SideWallImageData, $chart->BackWallImageData, SKColors::Transparent());
    }

    public static function DrawFloorGrid($canvas, $area, $chart, $wallDepth, $grid, $count)
    {
        $floorLeft = Column3DGeometry::GridPoint($area, 0.0, 0.0);
        $floorRight = Column3DGeometry::GridPoint($area, 1.0, 0.0);
        $frontLeft = self::Offset($floorLeft, $wallDepth);
        $frontRight = Column3DGeometry::UsesSeriesDepth($chart)
            ? Column3DGeometry::ResolveSeriesFloorPoint($area, 1.0, 1.0)
            : self::Offset($floorRight, $wallDepth);
        $canvas->DrawLine($floorLeft, $floorRight, $grid);
        $canvas->DrawLine($floorLeft, $frontLeft, $grid);
        $canvas->DrawLine($frontLeft, $frontRight, $grid);

        for ($index = 0; $index <= $count; $index++) {
            $ratio = $count > 0 ? $index / (float)$count : 0.0;
            $front = Column3DGeometry::UsesSeriesDepth($chart)
                ? Column3DGeometry::ResolveSeriesFloorPoint($area, $ratio, 1.0)
                : self::Offset(self::Interpolate($floorLeft, $floorRight, $ratio), $wallDepth);
            $canvas->DrawLine($front->X, $front->Y, $front->X, $front->Y + $area->Height * 0.012, $grid);
        }
    }

    private static function Offset($point, $offset)
    {
        return new SKPoint($point->X + $offset->X, $point->Y + $offset->Y);
    }

    private static function Interpolate($start, $end, $ratio)
    {
        return new SKPoint($start->X + ($end->X - $start->X) * (float)$ratio,
            $start->Y + ($end->Y - $start->Y) * (float)$ratio);
    }
}
