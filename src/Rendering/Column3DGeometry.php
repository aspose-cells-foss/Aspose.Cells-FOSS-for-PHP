<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;

final class Column3DGeometry
{
    const CylinderFloorExpansion = 1.35;

    const PercentStackedBoxFloorExpansion = 1.3;

    const SeriesWallLeftDepthXRatio = 0.196;

    const SeriesWallLeftDepthYRatio = 0.089;

    const SeriesWallRightDepthXRatio = 0.0165;

    const SeriesWallRightDepthYRatio = 0.158;

    private function __construct()
    {
    }

    public static function IsPercentStackedBox($chart)
    {
        return $chart !== null && $chart->IsPercentStacked
            && $chart->BarShape === 'box';
    }

    public static function UsesSeriesDepth($chart)
    {
        return $chart !== null && $chart->IsBar3D && !$chart->IsStacked
            && !$chart->IsPercentStacked && $chart->HasVisibleSeriesAxis;
    }

    public static function ResolveFloorDepth($area, $chart, $defaultDepth, $barWidth, $shapeDepth, $horizontalSlope)
    {
        if ($chart === null) {
            return $defaultDepth;
        }

        if (self::UsesSeriesDepth($chart)) {
            return self::ResolveSeriesWallDepth($area, 0.0);
        }

        if ($chart->BarShape === 'cylinder') {
            $widthAxisY = $barWidth * $horizontalSlope * 0.5;
            $depthAxisY = $shapeDepth->Y * 0.5;
            $verticalRadius = (float)sqrt(
                $widthAxisY * $widthAxisY + $depthAxisY * $depthAxisY);
            $floorHeight = (-$shapeDepth->Y + $verticalRadius * 2.0 + 0.5)
                * self::CylinderFloorExpansion;
            return new SKPoint(min($defaultDepth->X, -$floorHeight * 2.0),
                max($defaultDepth->Y, $floorHeight));
        }

        if (!self::IsPercentStackedBox($chart)) {
            return $defaultDepth;
        }

        $projectedWidthHeight = abs($barWidth * $horizontalSlope);
        $boxFloorHeight = (-$shapeDepth->Y + $projectedWidthHeight + 0.5)
            * self::PercentStackedBoxFloorExpansion;
        return new SKPoint(min($defaultDepth->X, -$boxFloorHeight * 2.0),
            max($defaultDepth->Y, $boxFloorHeight));
    }

    public static function ResolveBoxBottomCenterOffset($floorMidpoint, $depth)
    {
        return new SKPoint($floorMidpoint->X - $depth->X * 0.5,
            $floorMidpoint->Y - $depth->Y * 0.5);
    }

    public static function ResolveSeriesFloorOffset($floorDepth, $seriesIndex, $seriesCount)
    {
        $ratio = $seriesCount > 1
            ? 0.85 - 0.60 * $seriesIndex / ($seriesCount - 1)
            : 0.5;
        return new SKPoint($floorDepth->X * $ratio, $floorDepth->Y * $ratio);
    }

    public static function ResolveSeriesWallDepth($area, $horizontalRatio)
    {
        $left = new SKPoint(-$area->Width * self::SeriesWallLeftDepthXRatio,
            $area->Height * self::SeriesWallLeftDepthYRatio);
        $right = new SKPoint(-$area->Width * self::SeriesWallRightDepthXRatio,
            $area->Height * self::SeriesWallRightDepthYRatio);
        return self::Interpolate($left, $right, $horizontalRatio);
    }

    public static function ResolveSeriesFloorPoint($area, $horizontalRatio, $depthRatio)
    {
        $back = self::GridPoint($area, $horizontalRatio, 0.0);
        $depth = self::ResolveSeriesWallDepth($area, $horizontalRatio);
        return new SKPoint($back->X + $depth->X * (float)$depthRatio,
            $back->Y + $depth->Y * (float)$depthRatio);
    }

    public static function ResolveSeriesShapeDepth($area, $barWidth, $horizontalRatio)
    {
        $wallDepth = self::ResolveSeriesWallDepth($area, $horizontalRatio);
        $wallLength = (float)sqrt(
            $wallDepth->X * $wallDepth->X + $wallDepth->Y * $wallDepth->Y);
        if ($wallLength < 0.0001) {
            return new SKPoint($barWidth * 0.85, 0.0);
        }
        $scale = $barWidth * 0.85 / $wallLength;
        return new SKPoint(-$wallDepth->X * $scale, -$wallDepth->Y * $scale);
    }

    public static function GridPoint($area, $horizontalRatio, $valueRatio)
    {
        $bottomLeft = new SKPoint($area->Left + $area->Width * 0.248,
            $area->Top + $area->Height * 0.593);
        $bottomRight = new SKPoint($area->Left + $area->Width * 0.838,
            $area->Top + $area->Height * 0.738);
        $topLeft = new SKPoint($area->Left + $area->Width * 0.233,
            $area->Top + $area->Height * 0.193);
        $topRight = new SKPoint($area->Left + $area->Width * 0.860,
            $area->Top + $area->Height * 0.246);
        $bottom = self::Interpolate($bottomLeft, $bottomRight, $horizontalRatio);
        $top = self::Interpolate($topLeft, $topRight, $horizontalRatio);
        return self::Interpolate($bottom, $top, $valueRatio);
    }

    private static function Interpolate($start, $end, $ratio)
    {
        return new SKPoint(
            $start->X + ($end->X - $start->X) * (float)$ratio,
            $start->Y + ($end->Y - $start->Y) * (float)$ratio);
    }

    public static function ResolveBoxHeightExtension($bottomCenterOffset, $depth)
    {
        return max(0.0, $bottomCenterOffset->Y + $depth->Y - 2.5);
    }

    public static function ResolveShapeOffsets($chart, $barWidth, $depth, $horizontalSlope, &$floorOffset, &$heightExtension)
    {
        $heightExtension = 0.0;
        if ($chart->BarShape === 'cylinder') {
            $floorOffset = Column3DCylinderRenderer::ResolveBottomCenterOffset(
                $floorOffset, $depth);
            $heightExtension = Column3DCylinderRenderer::ResolveHeightExtension(
                $floorOffset->Y, $barWidth, $depth, $horizontalSlope);
        } elseif (self::IsPercentStackedBox($chart)) {
            $floorOffset = self::ResolveBoxBottomCenterOffset($floorOffset, $depth);
            $heightExtension = self::ResolveBoxHeightExtension($floorOffset, $depth);
        }
    }
}
