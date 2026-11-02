<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;

final class Column3DChartRenderer
{
    private const LabelSizePt = 9.0;

    private $context;

    private $dateSystem;

    private $culture;

    private $chartArea;

    public function __construct($context, $dateSystem, $culture)
    {
        $this->context = $context;
        $this->dateSystem = $dateSystem;
        $this->culture = $culture !== null ? $culture : CultureInfo::GetInvariantCulture();
    }

    public function Draw($canvas, $area, $chart, $min, $max, $major)
    {
        if ($canvas === null || $chart === null || count($chart->Series) === 0 || $max <= $min) {
            return;
        }

        $this->chartArea = $area;
        if (Column3DGeometry::UsesSeriesDepth($chart)
            && $chart->BarShape === 'box') {
            $scene = new Column3DSceneRenderer($this->context, $this->dateSystem, $this->culture);
            $scene->Draw($canvas, $area, $chart, $min, $max, $major);
            return;
        }

        $projectionArea = self::ResolveProjectionArea($area, $chart);
        self::DrawPlotBox($canvas, $projectionArea, $chart, $min, $max, $major);
        self::DrawColumns($canvas, $projectionArea, $chart, $min, $max);
        $this->DrawDataLabels($canvas, $projectionArea, $chart, $min, $max);
        $this->DrawValueLabels($canvas, $projectionArea, $chart, $min, $max, $major);
        $this->DrawCategoryLabels($canvas, $projectionArea, $chart);
    }

    private static function DrawPlotBox($canvas, $area, $chart, $min, $max, $major)
    {
        $grid = new SKPaint();
        $grid->Style = SKPaintStyle::Stroke;
        $grid->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $grid->StrokeWidth = 0.75;
        $grid->IsAntialias = true;

        $wallDepth = self::ResolveFloorDepth($area, $chart);
        Column3DWallRenderer::Draw($canvas, $area, $chart, $wallDepth);
        if ($chart->HasMinorValueGridlines) {
            $minor = ($chart->MinorUnit !== null && $chart->MinorUnit > 0.0)
                ? $chart->MinorUnit
                : $major / 5.0;
            if ($minor > 0.0) {
                $minorGrid = new SKPaint();
                $minorGrid->Style = SKPaintStyle::Stroke;
                $minorGrid->Color = $chart->MinorValueGridlineColor;
                $minorGrid->StrokeWidth = (float)max(0.25, $chart->MinorValueGridlineWidthPt);
                $minorGrid->IsAntialias = true;
                $minorSteps = (int)floor(($max - $min) / $minor + 0.0001);
                for ($index = 1; $index < $minorSteps; $index++) {
                    $value = $min + $index * $minor;
                    $majorRatio = ($value - $min) / $major;
                    if (abs($majorRatio - round($majorRatio)) < 0.0001) {
                        continue;
                    }
                    $ratio = ($value - $min) / ($max - $min);
                    $left = Column3DGeometry::GridPoint($area, 0.0, $ratio);
                    $right = Column3DGeometry::GridPoint($area, 1.0, $ratio);
                    $canvas->DrawLine(self::Offset($left, $wallDepth), $left, $minorGrid);
                    $canvas->DrawLine($left, $right, $minorGrid);
                }

                $minorGrid->Dispose();
            }
        }

        $steps = (int)round(($max - $min) / $major);
        for ($index = 0; $index <= $steps; $index++) {
            $ratio = $index * $major / ($max - $min);
            $left = Column3DGeometry::GridPoint($area, 0.0, $ratio);
            $right = Column3DGeometry::GridPoint($area, 1.0, $ratio);
            $canvas->DrawLine(self::Offset($left, $wallDepth), $left, $grid);
            $canvas->DrawLine($left, $right, $grid);
        }

        Column3DWallRenderer::DrawFloorGrid($canvas, $area, $chart, $wallDepth,
            $grid, self::ResolvePointCount($chart));

        if ($chart->HasValueAxisLine) {
            $axis = new SKPaint();
            $axis->Style = SKPaintStyle::Stroke;
            $axis->Color = $chart->ValueAxisLineColor;
            $axis->StrokeWidth = (float)max(0.25, $chart->ValueAxisLineWidthPt);
            $axis->IsAntialias = true;
            $axisBottom = self::ValueAxisPoint($area, 0.0, $chart);
            $axisTop = self::ValueAxisPoint($area, 1.0, $chart);
            $canvas->DrawLine($axisBottom, $axisTop, $axis);
            if ($chart->ValueAxisMajorTickMark === 'cross') {
                for ($index = 0; $index <= $steps; $index++) {
                    $ratio = $index * $major / ($max - $min);
                    $point = self::ValueAxisPoint($area, $ratio, $chart);
                    $canvas->DrawLine($point->X - 3.0, $point->Y, $point->X + 3.0, $point->Y, $axis);
                }
            }

            $axis->Dispose();
        }

        $grid->Dispose();
    }

    private static function ResolveFloorDepth($area, $chart)
    {
        $defaultDepth = new SKPoint(-$area->Width * 0.035, $area->Height * 0.014);
        $pointCount = self::ResolvePointCount($chart);
        $barWidth = $chart->IsStacked || $chart->IsPercentStacked
            ? self::ResolveStackedColumnWidth($area, $chart, $pointCount)
            : $area->Width * 0.022;
        $shapeDepth = self::ResolveShapeDepth($barWidth, $chart);
        $horizontalSlope = self::ResolveHorizontalAxisSlope($area, 0.0);
        return Column3DGeometry::ResolveFloorDepth($area, $chart, $defaultDepth,
            $barWidth, $shapeDepth, $horizontalSlope);
    }

    private static function ValueAxisPoint($area, $valueRatio, $chart)
    {
        return self::Offset(Column3DGeometry::GridPoint($area, 0.0, $valueRatio),
            self::ResolveFloorDepth($area, $chart));
    }

    private static function DrawColumns($canvas, $area, $chart, $min, $max)
    {
        if ($chart->IsStacked || $chart->IsPercentStacked) {
            self::DrawStackedColumns($canvas, $area, $chart, $min, $max);
            return;
        }

        self::DrawClusteredColumns($canvas, $area, $chart, $min, $max);
    }

    private static function DrawClusteredColumns($canvas, $area, $chart, $min, $max)
    {
        $pointCount = self::ResolvePointCount($chart);
        if ($pointCount <= 0) {
            return;
        }

        $usesSeriesDepth = Column3DGeometry::UsesSeriesDepth($chart);
        $barWidth = $area->Width * ($usesSeriesDepth ? 0.035 : 0.022);
        $seriesPitch = self::ResolveClusteredSeriesPitch($barWidth);
        $depth = self::ResolveShapeDepth($barWidth, $chart);
        $groupSpan = (count($chart->Series) - 1) * $seriesPitch;
        $horizontalAxisSlope = self::ResolveHorizontalAxisSlope($area, 0.0);
        $verticalAxisSlope = self::ResolveVerticalAxisSlope($area, $chart);
        $floorMidpointOffset = self::ResolveFloorMidpointOffset($area, $chart);
        $heightExtension = 0.0;
        Column3DGeometry::ResolveShapeOffsets($chart, $barWidth, $depth,
            $horizontalAxisSlope, $floorMidpointOffset, $heightExtension);

        for ($categoryIndex = 0; $categoryIndex < $pointCount; $categoryIndex++) {
            $categoryRatio = ($categoryIndex + 0.5) / $pointCount;
            $gridBaseline = Column3DGeometry::GridPoint($area, $categoryRatio, 0.0);
            $baseline = self::Offset($gridBaseline, $floorMidpointOffset);
            $categoryDepth = $usesSeriesDepth
                ? Column3DGeometry::ResolveSeriesShapeDepth($area, $barWidth, $categoryRatio)
                : $depth;
            $firstSeries = $usesSeriesDepth ? count($chart->Series) - 1 : 0;
            $lastSeries = $usesSeriesDepth ? -1 : count($chart->Series);
            $seriesStep = $usesSeriesDepth ? -1 : 1;
            for ($seriesIndex = $firstSeries; $seriesIndex != $lastSeries; $seriesIndex += $seriesStep) {
                $series = $chart->Series[$seriesIndex];
                if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                    continue;
                }

                $valueRatio = ($series->Values[$categoryIndex] - $min) / ($max - $min);
                if ($valueRatio < 0.0) {
                    $valueRatio = 0.0;
                } elseif ($valueRatio > 1.0) {
                    $valueRatio = 1.0;
                }

                $centerOffset = $seriesIndex * $seriesPitch - $groupSpan * 0.5;
                $projectedOffset = self::ProjectAlongHorizontalAxis($centerOffset, $horizontalAxisSlope);
                $bottomCenter = $usesSeriesDepth
                    ? self::Offset($gridBaseline, Column3DGeometry::ResolveSeriesFloorOffset(
                        Column3DGeometry::ResolveSeriesWallDepth($area, $categoryRatio),
                        $seriesIndex, count($chart->Series)))
                    : new SKPoint($baseline->X + $projectedOffset->X,
                        $baseline->Y + $projectedOffset->Y);
                $height = self::ValueHeight($area, $categoryRatio, $valueRatio)
                    + (float)$valueRatio * $heightExtension;
                $heightOffset = self::ProjectAlongVerticalAxis($height, $verticalAxisSlope);
                self::DrawColumnShape($canvas, $bottomCenter, $barWidth, $heightOffset, $categoryDepth,
                    $horizontalAxisSlope, $usesSeriesDepth ? self::ResolveHorizontalAxisSlope(
                        $area, $valueRatio) : $horizontalAxisSlope, $chart, $series, true);
            }
        }
    }

    private static function DrawStackedColumns($canvas, $area, $chart, $min, $max)
    {
        $pointCount = self::ResolvePointCount($chart);
        if ($pointCount <= 0) {
            return;
        }

        $barWidth = self::ResolveStackedColumnWidth($area, $chart, $pointCount);
        $depth = self::ResolveShapeDepth($barWidth, $chart);
        $horizontalAxisSlope = self::ResolveHorizontalAxisSlope($area, 0.0);
        $verticalAxisSlope = self::ResolveVerticalAxisSlope($area, $chart);
        $floorMidpointOffset = self::ResolveFloorMidpointOffset($area, $chart);
        $heightExtension = 0.0;
        Column3DGeometry::ResolveShapeOffsets($chart, $barWidth, $depth,
            $horizontalAxisSlope, $floorMidpointOffset, $heightExtension);

        for ($categoryIndex = 0; $categoryIndex < $pointCount; $categoryIndex++) {
            $categoryRatio = ($categoryIndex + 0.5) / $pointCount;
            $baseline = self::Offset(Column3DGeometry::GridPoint($area, $categoryRatio, 0.0),
                $floorMidpointOffset);
            for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
                $series = $chart->Series[$seriesIndex];
                if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                    continue;
                }

                $startValue = 0.0;
                $endValue = 0.0;
                self::ResolveStackedSegment($chart, $categoryIndex, $seriesIndex, $startValue, $endValue);
                if (abs($endValue - $startValue) < 0.0000001) {
                    continue;
                }

                $startRatio = self::ClampRatio(($startValue - $min) / ($max - $min));
                $endRatio = self::ClampRatio(($endValue - $min) / ($max - $min));
                $startHeight = self::ValueHeight($area, $categoryRatio, $startRatio)
                    + (float)$startRatio * $heightExtension;
                $endHeight = self::ValueHeight($area, $categoryRatio, $endRatio)
                    + (float)$endRatio * $heightExtension;
                $bottomCenter = self::Offset(
                    $baseline,
                    self::ProjectAlongVerticalAxis($startHeight, $verticalAxisSlope));
                $heightOffset = self::ProjectAlongVerticalAxis(
                    $endHeight - $startHeight,
                    $verticalAxisSlope);
                $drawTopFace = self::IsLastStackedSegment(
                    $chart,
                    $categoryIndex,
                    $seriesIndex,
                    $endValue - $startValue);
                $bottomFaceSlope = Column3DGeometry::IsPercentStackedBox($chart)
                    ? self::ResolveHorizontalAxisSlope($area, $startRatio) : $horizontalAxisSlope;
                $topFaceSlope = Column3DGeometry::IsPercentStackedBox($chart)
                    ? self::ResolveHorizontalAxisSlope($area, $endRatio) : $horizontalAxisSlope;
                self::DrawColumnShape($canvas, $bottomCenter, $barWidth, $heightOffset, $depth,
                    $bottomFaceSlope, $topFaceSlope, $chart, $series, $drawTopFace);
            }
        }
    }

    private static function ResolveStackedColumnWidth($area, $chart, $pointCount)
    {
        if ($pointCount <= 0) {
            return 0.0;
        }

        $left = Column3DGeometry::GridPoint($area, 0.0, 0.0);
        $right = Column3DGeometry::GridPoint($area, 1.0, 0.0);
        $categoryPitch = ($right->X - $left->X) / $pointCount;
        $gapRatio = max(0.0, $chart->GapWidthPercent / 100.0);
        return $categoryPitch / (float)(1.0 + $gapRatio);
    }

    private static function ResolveStackedSegment($chart, $categoryIndex, $seriesIndex, &$startValue, &$endValue)
    {
        $startValue = 0.0;
        $endValue = 0.0;
        if ($chart === null || $seriesIndex < 0 || $seriesIndex >= count($chart->Series)) {
            return;
        }

        $targetSeries = $chart->Series[$seriesIndex];
        if ($categoryIndex < 0 || $categoryIndex >= count($targetSeries->Values)
            || $targetSeries->Values[$categoryIndex] === null) {
            return;
        }

        $targetValue = self::NormalizeStackedValue(
            $chart,
            $categoryIndex,
            $targetSeries->Values[$categoryIndex]);
        $positive = $targetValue >= 0.0;
        for ($index = 0; $index < $seriesIndex; $index++) {
            $series = $chart->Series[$index];
            if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                continue;
            }

            $value = self::NormalizeStackedValue($chart, $categoryIndex, $series->Values[$categoryIndex]);
            if (($value >= 0.0) == $positive) {
                $startValue += $value;
            }
        }

        $endValue = $startValue + $targetValue;
    }

    private static function NormalizeStackedValue($chart, $categoryIndex, $value)
    {
        if (!$chart->IsPercentStacked) {
            return $value;
        }

        $total = 0.0;
        $positive = $value >= 0.0;
        for ($index = 0; $index < count($chart->Series); $index++) {
            $series = $chart->Series[$index];
            if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                continue;
            }

            $candidate = $series->Values[$categoryIndex];
            if (($candidate >= 0.0) == $positive) {
                $total += abs($candidate);
            }
        }

        return $total > 0.0 ? $value / $total : 0.0;
    }

    private static function IsLastStackedSegment($chart, $categoryIndex, $seriesIndex, $value)
    {
        $positive = $value >= 0.0;
        for ($index = $seriesIndex + 1; $index < count($chart->Series); $index++) {
            $series = $chart->Series[$index];
            if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                continue;
            }

            $laterValue = self::NormalizeStackedValue(
                $chart,
                $categoryIndex,
                $series->Values[$categoryIndex]);
            if (($laterValue >= 0.0) == $positive && abs($laterValue) > 0.0000001) {
                return false;
            }
        }

        return true;
    }

    private static function ClampRatio($value)
    {
        return max(0.0, min(1.0, $value));
    }

    private static function ResolveFloorMidpointOffset($area, $chart)
    {
        $floorDepth = self::ResolveFloorDepth($area, $chart);
        return new SKPoint($floorDepth->X * 0.5, $floorDepth->Y * 0.5);
    }

    private static function ResolveClusteredSeriesPitch($barWidth)
    {
        return $barWidth;
    }

    private static function ResolveColumnDepth($barWidth)
    {
        return new SKPoint($barWidth * 0.55, -$barWidth * 0.24);
    }

    private static function ResolveShapeDepth($barWidth, $chart)
    {
        if (Column3DGeometry::IsPercentStackedBox($chart)) {
            return new SKPoint($barWidth * 0.55, -$barWidth * 0.275);
        }

        return self::ResolveColumnDepth($barWidth);
    }

    private static function ResolveProjectionArea($area, $chart)
    {
        if (Column3DGeometry::UsesSeriesDepth($chart)) {
            return new SKRect(
                $area->Left + $area->Width * 0.13,
                $area->Top + $area->Height * 0.08,
                $area->Right - $area->Width * 0.04,
                $area->Bottom - $area->Height * 0.17);
        }

        if (Column3DGeometry::IsPercentStackedBox($chart)) {
            return new SKRect($area->Left, $area->Top, $area->Right,
                $area->Bottom - $area->Height * 0.043);
        }

        if ($chart === null || $chart->BarShape !== 'cylinder') {
            return $area;
        }

        return new SKRect(
            $area->Left + $area->Width * 0.12,
            $area->Top + $area->Height * 0.10,
            $area->Right - $area->Width * 0.18,
            $area->Bottom - $area->Height * 0.20);
    }

    private static function ResolveHorizontalAxisSlope($area, $valueRatio)
    {
        $left = Column3DGeometry::GridPoint($area, 0.0, $valueRatio);
        $right = Column3DGeometry::GridPoint($area, 1.0, $valueRatio);
        $width = $right->X - $left->X;
        if (abs($width) < 0.0001) {
            return 0.0;
        }

        return ($right->Y - $left->Y) / $width;
    }

    private static function ProjectAlongHorizontalAxis($distance, $slope)
    {
        return new SKPoint($distance, $distance * $slope);
    }

    private static function ResolveVerticalAxisSlope($area, $chart)
    {
        $bottom = self::ValueAxisPoint($area, 0.0, $chart);
        $top = self::ValueAxisPoint($area, 1.0, $chart);
        $height = $bottom->Y - $top->Y;
        if (abs($height) < 0.0001) {
            return 0.0;
        }

        return ($top->X - $bottom->X) / $height;
    }

    private static function ProjectAlongVerticalAxis($height, $slope)
    {
        return new SKPoint($height * $slope, -$height);
    }

    private static function DrawColumnShape($canvas, $bottomCenter, $width, $heightOffset, $depth, $bottomAxisSlope, $topAxisSlope, $chart, $series, $drawTopFace)
    {
        if ($chart->BarShape === 'cylinder') {
            Column3DCylinderRenderer::Draw($canvas, $bottomCenter, $width, $heightOffset,
                $depth, $bottomAxisSlope, $series, $drawTopFace);
            return;
        }

        self::DrawPrism($canvas, $bottomCenter, $width, $heightOffset, $depth,
            $bottomAxisSlope, $topAxisSlope, $series, $drawTopFace);
    }

    private static function DrawPrism($canvas, $bottomCenter, $width, $heightOffset, $depth, $bottomAxisSlope, $topAxisSlope, $series, $drawTopFace)
    {
        $color = $series->Color;
        $halfWidth = $width * 0.5;
        $bottomSlopeOffset = $width * $bottomAxisSlope * 0.5;
        $topSlopeOffset = $width * $topAxisSlope * 0.5;
        $frontBottomLeft = new SKPoint(
            $bottomCenter->X - $halfWidth, $bottomCenter->Y - $bottomSlopeOffset);
        $frontBottomRight = new SKPoint(
            $bottomCenter->X + $halfWidth, $bottomCenter->Y + $bottomSlopeOffset);
        $topCenter = self::Offset($bottomCenter, $heightOffset);
        $frontTopLeft = new SKPoint($topCenter->X - $halfWidth, $topCenter->Y - $topSlopeOffset);
        $frontTopRight = new SKPoint($topCenter->X + $halfWidth, $topCenter->Y + $topSlopeOffset);
        $backTopLeft = self::Offset($frontTopLeft, $depth);
        $backTopRight = self::Offset($frontTopRight, $depth);
        $backBottomRight = self::Offset($frontBottomRight, $depth);

        $frontFill = new SKPaint();
        $frontFill->Style = SKPaintStyle::Fill;
        $frontFill->Color = $color;
        $frontFill->IsAntialias = true;
        $topFill = new SKPaint();
        $topFill->Style = SKPaintStyle::Fill;
        $topFill->Color = self::Shade($color, 1.08);
        $topFill->IsAntialias = true;
        $sideFill = new SKPaint();
        $sideFill->Style = SKPaintStyle::Fill;
        $sideFill->Color = self::Shade($color, 0.72);
        $sideFill->IsAntialias = true;
        $edge = new SKPaint();
        $edge->Style = SKPaintStyle::Stroke;
        $edge->Color = $series->LineColor;
        $edge->StrokeWidth = (float)max(0.25, $series->LineWidthPt);
        $edge->IsAntialias = true;

        self::DrawQuad($canvas, $frontTopRight, $backTopRight, $backBottomRight, $frontBottomRight,
            $sideFill, $edge, $series->HasVisibleLine);
        if ($drawTopFace) {
            self::DrawQuad($canvas, $frontTopLeft, $backTopLeft, $backTopRight, $frontTopRight,
                $topFill, $edge, $series->HasVisibleLine);
        }

        $front = new SKPath();
        $front->MoveTo($frontTopLeft);
        $front->LineTo($frontTopRight);
        $front->LineTo($frontBottomRight);
        $front->LineTo($frontBottomLeft);
        $front->Close();
        $canvas->DrawPath($front, $frontFill);
        if ($series->HasVisibleLine) {
            $canvas->DrawPath($front, $edge);
        }

        $front->Dispose();
        $edge->Dispose();
        $sideFill->Dispose();
        $topFill->Dispose();
        $frontFill->Dispose();
    }

    private function DrawDataLabels($canvas, $area, $chart, $min, $max)
    {
        $pointCount = self::ResolvePointCount($chart);
        if ($pointCount <= 0 || $max <= $min) {
            return;
        }

        $barWidth = $area->Width * 0.022;
        $seriesPitch = self::ResolveClusteredSeriesPitch($barWidth);
        $groupSpan = (count($chart->Series) - 1) * $seriesPitch;
        $horizontalAxisSlope = self::ResolveHorizontalAxisSlope($area, 0.0);
        $verticalAxisSlope = self::ResolveVerticalAxisSlope($area, $chart);
        $floorMidpointOffset = self::ResolveFloorMidpointOffset($area, $chart);

        for ($categoryIndex = 0; $categoryIndex < $pointCount; $categoryIndex++) {
            $categoryRatio = ($categoryIndex + 0.5) / $pointCount;
            $baseline = self::Offset(Column3DGeometry::GridPoint($area, $categoryRatio, 0.0),
                $floorMidpointOffset);
            for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
                $series = $chart->Series[$seriesIndex];
                if (!$series->ShowValueDataLabels || $categoryIndex >= count($series->Values)
                    || $series->Values[$categoryIndex] === null) {
                    continue;
                }

                $value = $series->Values[$categoryIndex];
                $valueRatio = max(0.0, min(1.0, ($value - $min) / ($max - $min)));
                $centerOffset = $seriesIndex * $seriesPitch - $groupSpan * 0.5;
                $projectedOffset = self::ProjectAlongHorizontalAxis($centerOffset, $horizontalAxisSlope);
                $bottomCenter = new SKPoint(
                    $baseline->X + $projectedOffset->X,
                    $baseline->Y + $projectedOffset->Y);
                $height = self::ValueHeight($area, $categoryRatio, $valueRatio);
                $heightOffset = self::ProjectAlongVerticalAxis($height, $verticalAxisSlope);
                $labelX = $bottomCenter->X + $heightOffset->X;
                $labelY = $bottomCenter->Y + $heightOffset->Y - 6.0;
                if ($series->DataLabelPosition === 'ctr') {
                    $labelX = $bottomCenter->X + $heightOffset->X * 0.5;
                    $labelY = $bottomCenter->Y + $heightOffset->Y * 0.5;
                }

                $formatCode = !($series->ValueFormatCode === null || $series->ValueFormatCode === '')
                    ? $series->ValueFormatCode
                    : $chart->ValueFormatCode;
                $label = ChartXmlParser::FormatLabel($value, $formatCode, $this->dateSystem, $this->culture);
                $this->DrawText($canvas, $area, $label !== null ? rtrim($label) : '',
                    $labelX, $labelY, $series->DataLabelColor, false);
            }
        }
    }

    private static function DrawQuad($canvas, $a, $b, $c, $d, $fill, $edge, $drawEdge)
    {
        $path = new SKPath();
        $path->MoveTo($a);
        $path->LineTo($b);
        $path->LineTo($c);
        $path->LineTo($d);
        $path->Close();
        $canvas->DrawPath($path, $fill);
        if ($drawEdge) {
            $canvas->DrawPath($path, $edge);
        }

        $path->Dispose();
    }

    private function DrawValueLabels($canvas, $area, $chart, $min, $max, $major)
    {
        $steps = (int)round(($max - $min) / $major);
        for ($index = 0; $index <= $steps; $index++) {
            $value = $min + $index * $major;
            $ratio = ($value - $min) / ($max - $min);
            $point = self::ValueAxisPoint($area, $ratio, $chart);
            $label = ChartXmlParser::FormatLabel($value, $chart->ValueFormatCode,
                $this->dateSystem, $this->culture);
            $this->DrawText($canvas, $area, $label !== null ? rtrim($label) : '',
                $point->X - 8.0, $point->Y, $chart->AxisLabelColor, true);
        }
    }

    private function DrawCategoryLabels($canvas, $area, $chart)
    {
        $count = self::ResolvePointCount($chart);
        for ($index = 0; $index < $count; $index++) {
            $ratio = ($index + 0.5) / $count;
            $point = Column3DGeometry::UsesSeriesDepth($chart)
                ? Column3DGeometry::ResolveSeriesFloorPoint($area, $ratio, 1.0)
                : self::Offset(Column3DGeometry::GridPoint($area, $ratio, 0.0),
                    self::ResolveFloorDepth($area, $chart));
            $label = $index < count($chart->Categories) ? $chart->Categories[$index] : '';
            if ($chart->BarShape === 'cylinder') {
                $this->DrawRotatedText($canvas, $label, $point->X, $point->Y + 15.0,
                    $chart->AxisLabelColor, -45.0);
            } else {
                $this->DrawText($canvas, $area, $label, $point->X, $point->Y + 13.0,
                    $chart->AxisLabelColor, false);
            }
        }
    }

    private function DrawRotatedText($canvas, $text, $x, $y, $color, $angleDegrees)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = self::LabelSizePt;
        $fontContext = $this->context->GetFontContext($fontValue);
        $runs = $fontContext->SplitRuns($text);
        if (count($runs) == 0) {
            return;
        }

        $metricsFont = new SKFont($runs[0]->Typeface, self::LabelSizePt);
        $metrics = $metricsFont->Metrics;
        $baseline = -($metrics->Ascent + $metrics->Descent) * 0.5;
        $cursor = -$fontContext->Measure($text) * 0.5;
        $canvas->Save();
        $canvas->ClipRect($this->chartArea);
        $canvas->Translate($x, $y);
        $canvas->RotateDegrees($angleDegrees);
        for ($index = 0; $index < count($runs); $index++) {
            $font = new SKFont($runs[$index]->Typeface, self::LabelSizePt);
            $font->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $runs[$index]->Text,
                $cursor, $baseline, $font, $color);
            $font->Dispose();
            $cursor += $runs[$index]->WidthPt;
        }

        $canvas->Restore();
        $metricsFont->Dispose();
    }

    private function DrawText($canvas, $clip, $text, $x, $y, $color, $alignRight)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = self::LabelSizePt;
        $fontContext = $this->context->GetFontContext($fontValue);
        $drawX = $alignRight ? $x - $fontContext->Measure($text) : $x - $fontContext->Measure($text) * 0.5;
        $runs = $fontContext->SplitRuns($text);
        if (count($runs) == 0) {
            return;
        }

        $metricsFont = new SKFont($runs[0]->Typeface, self::LabelSizePt);
        $metrics = $metricsFont->Metrics;
        $baseline = $y - ($metrics->Ascent + $metrics->Descent) * 0.5;
        $canvas->Save();
        $textClip = ($this->chartArea !== null && $this->chartArea->Width > 0.0 && $this->chartArea->Height > 0.0)
            ? $this->chartArea
            : $clip;
        $canvas->ClipRect($textClip);
        $cursor = $drawX;
        for ($index = 0; $index < count($runs); $index++) {
            $font = new SKFont($runs[$index]->Typeface, self::LabelSizePt);
            $font->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $runs[$index]->Text, $cursor, $baseline, $font, $color);
            $font->Dispose();
            $cursor += $runs[$index]->WidthPt;
        }

        $canvas->Restore();
        $metricsFont->Dispose();
    }

    private static function ValueHeight($area, $horizontalRatio, $valueRatio)
    {
        $bottom = Column3DGeometry::GridPoint($area, $horizontalRatio, 0.0);
        $top = Column3DGeometry::GridPoint($area, $horizontalRatio, $valueRatio);
        return $bottom->Y - $top->Y;
    }

    private static function ResolvePointCount($chart)
    {
        $count = max(count($chart->Categories), count($chart->CategoryValues));
        for ($index = 0; $index < count($chart->Series); $index++) {
            if (count($chart->Series[$index]->Values) > $count) {
                $count = count($chart->Series[$index]->Values);
            }
        }
        return $count;
    }

    private static function Interpolate($start, $end, $ratio)
    {
        return new SKPoint(
            $start->X + ($end->X - $start->X) * (float)$ratio,
            $start->Y + ($end->Y - $start->Y) * (float)$ratio);
    }

    private static function Offset($point, $offset)
    {
        return new SKPoint($point->X + $offset->X, $point->Y + $offset->Y);
    }

    private static function Shade($color, $factor)
    {
        $red = max(0.0, min(255.0, $color->Red * $factor));
        $green = max(0.0, min(255.0, $color->Green * $factor));
        $blue = max(0.0, min(255.0, $color->Blue * $factor));
        return new SKColor((int)round($red), (int)round($green), (int)round($blue), $color->Alpha);
    }
}
