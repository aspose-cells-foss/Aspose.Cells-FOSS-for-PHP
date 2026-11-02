<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKStrokeJoin;
use Aspose\Cells_FOSS\Skia\SKTextAlign;

// Renders area3DChart series as depth-separated vertical area planes with a projected plot box.
final class Area3DChartRenderer
{
    private const LabelSizePt = 9.0;

    private const MajorGridlineWidthPt = 0.5;

    private const ValueAxisWidthPt = 0.5;

    private const FrontAxisGap = 0.06;

    private const StackedFrontAxisGap = 0.28;

    private $context;

    private $dateSystem;

    private $culture;

    public function __construct($context, $dateSystem, $culture)
    {
        $this->context = $context;
        $this->dateSystem = $dateSystem;
        $this->culture = $culture !== null ? $culture : CultureInfo::GetInvariantCulture();
    }

    public function Draw($canvas, $area, $plot, $chart, $min, $max, $major)
    {
        if ($canvas === null || $chart === null || $chart->Series === null || count($chart->Series) === 0 || $max <= $min) {
            return;
        }

        $renderedMajor = self::ResolveMajorUnit($chart, $min, $max, $major);
        self::DrawPlotBox($canvas, $plot, $chart, $min, $max, $renderedMajor, $this->dateSystem);
        self::DrawSeries($canvas, $plot, $chart, $min, $max, $this->dateSystem);
        $this->DrawValueLabels($canvas, $area, $plot, $chart, $min, $max, $renderedMajor);
        $this->DrawCategoryLabels($canvas, $area, $plot, $chart);
        $this->DrawSeriesAxisLabels($canvas, $area, $plot, $chart);
    }

    private static function DrawPlotBox($canvas, $plot, $chart, $min, $max, $major, $dateSystem)
    {
        $frontLeft = self::Project($plot, $chart, 0.0, 0.0, 0.0);
        $frontRight = self::Project($plot, $chart, 1.0, 0.0, 0.0);
        $frontTopLeft = self::Project($plot, $chart, 0.0, 1.0, 0.0);
        $backLeft = self::Project($plot, $chart, 0.0, 0.0, 1.0);
        $backRight = self::Project($plot, $chart, 1.0, 0.0, 1.0);
        $backTopLeft = self::Project($plot, $chart, 0.0, 1.0, 1.0);
        $backTopRight = self::Project($plot, $chart, 1.0, 1.0, 1.0);

        $wallColor = $chart->HasChartAreaGradient ? SKColors::Transparent() : SKColors::White();
        $floorColor = $chart->HasChartAreaGradient ? SKColors::Transparent() : new SKColor(0xF8, 0xF8, 0xF8);
        $floorFill = new SKPaint();
        $floorFill->Style = SKPaintStyle::Fill;
        $floorFill->Color = $floorColor;
        $floorFill->IsAntialias = true;
        $border = new SKPaint();
        $border->Style = SKPaintStyle::Stroke;
        $border->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $border->StrokeWidth = self::MajorGridlineWidthPt;
        $border->IsAntialias = true;

        Area3DWallFillRenderer::Draw($canvas, $frontLeft, $frontTopLeft, $backLeft, $backTopLeft,
            $backRight, $backTopRight, $chart->SideWallImageData, $chart->BackWallImageData, $wallColor);
        $canvas->DrawLine($backTopLeft, $backTopRight, $border);
        $canvas->DrawLine($backLeft, $backRight, $border);

        $floor = new SKPath();
        $floor->MoveTo($backLeft);
        $floor->LineTo($backRight);
        $floor->LineTo($frontRight);
        $floor->LineTo($frontLeft);
        $floor->Close();
        $canvas->DrawPath($floor, $floorFill);
        if ($chart->IsStacked || $chart->IsPercentStacked) {
            $canvas->DrawLine($backLeft, $backRight, $border);
            $canvas->DrawLine($frontLeft, $frontRight, $border);
            $canvas->DrawLine($frontLeft, $backLeft, $border);
        } else {
            $canvas->DrawPath($floor, $border);
        }

        $floor->Dispose();

        if ($chart->HasMinorValueGridlines) {
            $minor = ($chart->MinorUnit !== null && $chart->MinorUnit > 0.0)
                ? $chart->MinorUnit
                : $major / 5.0;
            $minorGrid = new SKPaint();
            $minorGrid->Style = SKPaintStyle::Stroke;
            $minorGrid->Color = $chart->MinorValueGridlineColor;
            $minorGrid->StrokeWidth = (float)max(0.5, $chart->MinorValueGridlineWidthPt);
            $minorGrid->IsAntialias = true;

            $epsilon = $minor * 0.1;
            for ($value = $min + $minor; $value < $max - $epsilon; $value += $minor) {
                $majorRatio = ($value - $min) / $major;
                if (abs($majorRatio - round($majorRatio)) <= 0.001) {
                    continue;
                }

                self::DrawProjectedValueGridline($canvas, $plot, $chart, ($value - $min) / ($max - $min), $minorGrid);
            }

            $minorGrid->Dispose();
        }

        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $normalized = ($value - $min) / ($max - $min);
            self::DrawProjectedValueGridline($canvas, $plot, $chart, $normalized, $border);
        }

        $monthlyTicks = Area3DDateAxisMath::BuildMonthlyTicks($chart, $dateSystem, $chart->Culture);
        $categoryCount = count($monthlyTicks) > 0 ? count($monthlyTicks) : self::ResolvePointCount($chart);
        for ($i = 0; $i < $categoryCount; $i++) {
            $x = count($monthlyTicks) > 0 ? $monthlyTicks[$i][0] : Area3DDateAxisMath::CategoryPosition($chart, $i, $categoryCount, $dateSystem);
            $floorFront = self::Project($plot, $chart, $x, 0.0, 0.0);
            $floorBack = self::Project($plot, $chart, $x, 0.0, 1.0);
            if (!$chart->IsStacked && !$chart->IsPercentStacked) {
                $canvas->DrawLine($floorFront, $floorBack, $border);
            } else {
                self::DrawCategoryAxisTick($canvas, $floorFront, $border);
            }
        }

        self::DrawValueAxisLine($canvas, $plot, $chart, $min, $max, $major);

        $border->Dispose();
        $floorFill->Dispose();
    }

    private static function DrawCategoryAxisTick($canvas, $axisPoint, $paint)
    {
        $end = new SKPoint($axisPoint->X, $axisPoint->Y + 2.5);
        $canvas->DrawLine($axisPoint, $end, $paint);
    }

    private static function DrawValueAxisLine($canvas, $plot, $chart, $min, $max, $major)
    {
        if (!$chart->HasValueAxisLine) {
            return;
        }

        $axis = new SKPaint();
        $axis->Style = SKPaintStyle::Stroke;
        $axis->Color = $chart->ValueAxisLineColor;
        $axis->StrokeWidth = self::ValueAxisWidthPt;
        $axis->IsAntialias = true;

        $axisX = $chart->ValueAxisOnRight ? 1.0 : 0.0;
        $axisDepth = $chart->ValueAxisOnRight ? 1.0 : 0.0;
        $bottom = self::Project($plot, $chart, $axisX, 0.0, $axisDepth);
        $top = self::Project($plot, $chart, $axisX, 1.0, $axisDepth);
        $canvas->DrawLine($bottom, $top, $axis);

        if ($chart->ValueAxisMajorTickMark === 'none') {
            $axis->Dispose();
            return;
        }

        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $point = self::Project($plot, $chart, $axisX, ($i * $major) / ($max - $min), $axisDepth);
            $outward = $chart->ValueAxisOnRight ? 1.0 : -1.0;
            $tickStart = $point->X;
            $tickEnd = $point->X + 3.0 * $outward;
            if ($chart->ValueAxisMajorTickMark === 'cross') {
                $tickStart = $point->X - 3.0;
                $tickEnd = $point->X + 3.0;
            } elseif ($chart->ValueAxisMajorTickMark === 'in') {
                $tickEnd = $point->X - 3.0 * $outward;
            }

            $canvas->DrawLine($tickStart, $point->Y, $tickEnd, $point->Y, $axis);
        }

        $axis->Dispose();
    }

    private static function DrawProjectedValueGridline($canvas, $plot, $chart, $normalized, $paint)
    {
        $atFrontLeft = self::Project($plot, $chart, 0.0, $normalized, 0.0);
        $atBackLeft = self::Project($plot, $chart, 0.0, $normalized, 1.0);
        $atBackRight = self::Project($plot, $chart, 1.0, $normalized, 1.0);
        $line = new SKPath();
        $line->MoveTo($atFrontLeft);
        $line->LineTo($atBackLeft);
        $line->LineTo($atBackRight);
        $canvas->DrawPath($line, $paint);
        $line->Dispose();
    }

    private static function DrawSeries($canvas, $plot, $chart, $min, $max, $dateSystem)
    {
        $pointCount = self::ResolvePointCount($chart);
        if ($pointCount <= 0) {
            return;
        }

        if ($chart->IsStacked || $chart->IsPercentStacked) {
            self::DrawStackedSeries($canvas, $plot, $chart, $min, $max, $pointCount, $dateSystem);
            return;
        }

        for ($seriesIndex = count($chart->Series) - 1; $seriesIndex >= 0; $seriesIndex--) {
            $series = $chart->Series[$seriesIndex];
            $values = ChartStackingMath::BuildDisplayValues($chart, $seriesIndex);
            $thickness = self::ResolveSeriesThickness(count($chart->Series));
            $frontDepth = self::SeriesFrontDepth($seriesIndex, count($chart->Series), $thickness);
            $backDepth = $frontDepth + $thickness;
            $frontPoints = array();
            $backPoints = array();
            $xPositions = array();
            for ($i = 0; $i < count($values); $i++) {
                if ($values[$i] === null) {
                    continue;
                }

                $x = Area3DDateAxisMath::CategoryPosition($chart, $i, $pointCount, $dateSystem);
                $y = self::NormalizeValue($values[$i], $min, $max);
                $xPositions[] = $x;
                $frontPoints[] = self::Project($plot, $chart, $x, $y, $frontDepth);
                $backPoints[] = self::Project($plot, $chart, $x, $y, $backDepth);
            }

            if (count($frontPoints) === 0) {
                continue;
            }

            Area3DDateAxisMath::SortByPosition($xPositions, $frontPoints, $backPoints);
            $baseline = self::NormalizeValue(max($min, 0.0), $min, $max);
            $firstX = $xPositions[0];
            $lastX = $xPositions[count($xPositions) - 1];
            $frontFirstBase = self::Project($plot, $chart, $firstX, $baseline, $frontDepth);
            $frontLastBase = self::Project($plot, $chart, $lastX, $baseline, $frontDepth);
            $backFirstBase = self::Project($plot, $chart, $firstX, $baseline, $backDepth);
            $backLastBase = self::Project($plot, $chart, $lastX, $baseline, $backDepth);

            self::DrawAreaFace($canvas, $backPoints, $backFirstBase, $backLastBase, self::Shade($series->Color, 0.84), self::Shade($series->Color, 0.65));
            $topFill = new SKPaint();
            $topFill->Style = SKPaintStyle::Fill;
            $topFill->Color = self::Shade($series->Color, 0.92);
            $topFill->IsAntialias = true;
            $sideFill = new SKPaint();
            $sideFill->Style = SKPaintStyle::Fill;
            $sideFill->Color = self::Shade($series->Color, 0.70);
            $sideFill->IsAntialias = true;

            for ($pointIndex = 0; $pointIndex < count($frontPoints) - 1; $pointIndex++) {
                self::DrawQuad($canvas, $frontPoints[$pointIndex], $frontPoints[$pointIndex + 1], $backPoints[$pointIndex + 1], $backPoints[$pointIndex], $topFill);
            }

            self::DrawQuad($canvas, $frontPoints[0], $backPoints[0], $backFirstBase, $frontFirstBase, $sideFill);
            self::DrawQuad($canvas, $frontPoints[count($frontPoints) - 1], $backPoints[count($backPoints) - 1], $backLastBase, $frontLastBase, $sideFill);

            $sideFill->Dispose();
            $topFill->Dispose();

            self::DrawAreaFace($canvas, $frontPoints, $frontFirstBase, $frontLastBase, $series->Color, self::Shade($series->Color, 0.72));
        }
    }

    private static function DrawAreaFace($canvas, $points, $firstBase, $lastBase, $fillColor, $edgeColor)
    {
        $path = new SKPath();
        $path->MoveTo($points[0]);
        for ($i = 1; $i < count($points); $i++) {
            $path->LineTo($points[$i]);
        }

        $path->LineTo($lastBase);
        $path->LineTo($firstBase);
        $path->Close();
        $fill = new SKPaint();
        $fill->Style = SKPaintStyle::Fill;
        $fill->Color = $fillColor;
        $fill->IsAntialias = true;
        $edge = new SKPaint();
        $edge->Style = SKPaintStyle::Stroke;
        $edge->Color = $edgeColor;
        $edge->StrokeWidth = 1.0;
        $edge->IsAntialias = true;
        $edge->StrokeJoin = SKStrokeJoin::Round;
        $canvas->DrawPath($path, $fill);
        $canvas->DrawPath($path, $edge);
        $edge->Dispose();
        $fill->Dispose();
        $path->Dispose();
    }

    private static function DrawStackedSeries($canvas, $plot, $chart, $min, $max, $pointCount, $dateSystem)
    {
        $thickness = self::ResolveStackedSeriesThickness($chart);
        $frontDepth = self::StackedFrontAxisGap;
        $backDepth = $frontDepth + $thickness;
        $baseline = self::NormalizeValue(max($min, 0.0), $min, $max);
        $frontUppers = array();
        $frontLowers = array();
        $backUppers = array();
        $backLowers = array();

        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $upperValues = ChartStackingMath::BuildDisplayValues($chart, $seriesIndex);
            $lowerValues = $seriesIndex > 0
                ? ChartStackingMath::BuildDisplayValues($chart, $seriesIndex - 1)
                : null;
            $frontUpper = array();
            $frontLower = array();
            $backUpper = array();
            $backLower = array();

            for ($i = 0; $i < count($upperValues); $i++) {
                if ($upperValues[$i] === null) {
                    continue;
                }

                $lower = $baseline;
                if ($lowerValues !== null && $i < count($lowerValues) && $lowerValues[$i] !== null) {
                    $lower = self::NormalizeValue($lowerValues[$i], $min, $max);
                }

                $x = Area3DDateAxisMath::CategoryPosition($chart, $i, $pointCount, $dateSystem);
                $upper = self::NormalizeValue($upperValues[$i], $min, $max);
                $frontUpper[] = self::ProjectStackedPoint($plot, $chart, $x, $upper, $frontDepth);
                $frontLower[] = self::ProjectStackedPoint($plot, $chart, $x, $lower, $frontDepth);
                $backUpper[] = self::ProjectStackedPoint($plot, $chart, $x, $upper, $backDepth);
                $backLower[] = self::ProjectStackedPoint($plot, $chart, $x, $lower, $backDepth);
            }

            $frontUppers[] = $frontUpper;
            $frontLowers[] = $frontLower;
            $backUppers[] = $backUpper;
            $backLowers[] = $backLower;
        }

        for ($seriesIndex = count($chart->Series) - 1; $seriesIndex >= 0; $seriesIndex--) {
            if (count($backUppers[$seriesIndex]) == 0) {
                continue;
            }

            $series = $chart->Series[$seriesIndex];
            self::DrawBandFace($canvas, $backUppers[$seriesIndex], $backLowers[$seriesIndex], $series, self::Shade($series->Color, 0.84), false);
        }

        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $frontUpper = $frontUppers[$seriesIndex];
            if (count($frontUpper) == 0) {
                continue;
            }

            $frontLower = $frontLowers[$seriesIndex];
            $backUpper = $backUppers[$seriesIndex];
            $backLower = $backLowers[$seriesIndex];
            $series = $chart->Series[$seriesIndex];
            $sideFill = new SKPaint();
            $sideFill->Style = SKPaintStyle::Fill;
            $sideFill->Color = self::Shade($series->Color, 0.70);
            $sideFill->IsAntialias = true;
            self::DrawQuad($canvas, $frontUpper[0], $backUpper[0], $backLower[0], $frontLower[0], $sideFill);
            $last = count($frontUpper) - 1;
            self::DrawQuad($canvas, $frontUpper[$last], $backUpper[$last], $backLower[$last], $frontLower[$last], $sideFill);
            $sideFill->Dispose();
        }

        $topSeriesIndex = count($chart->Series) - 1;
        if ($topSeriesIndex >= 0 && count($frontUppers[$topSeriesIndex]) > 0) {
            $frontUpper = $frontUppers[$topSeriesIndex];
            $backUpper = $backUppers[$topSeriesIndex];
            $series = $chart->Series[$topSeriesIndex];
            $topShade = $chart->IsPercentStacked ? 0.78 : 0.92;
            $topFill = new SKPaint();
            $topFill->Style = SKPaintStyle::Fill;
            $topFill->Color = self::Shade($series->Color, $topShade);
            $topFill->IsAntialias = true;
            for ($pointIndex = 0; $pointIndex < count($frontUpper) - 1; $pointIndex++) {
                self::DrawQuad($canvas, $frontUpper[$pointIndex], $frontUpper[$pointIndex + 1], $backUpper[$pointIndex + 1], $backUpper[$pointIndex], $topFill);
            }

            $topFill->Dispose();
        }

        for ($seriesIndex = count($chart->Series) - 1; $seriesIndex >= 0; $seriesIndex--) {
            if (count($frontUppers[$seriesIndex]) == 0) {
                continue;
            }

            $series = $chart->Series[$seriesIndex];
            self::DrawBandFace($canvas, $frontUppers[$seriesIndex], $frontLowers[$seriesIndex], $series, $series->Color, true);
        }
    }

    private static function ProjectStackedPoint($plot, $chart, $x, $y, $depth)
    {
        $point = self::Project($plot, $chart, $x, $y, $depth);
        if (!$chart->IsPercentStacked) {
            return $point;
        }

        if (abs($y - 1.0) > 0.000001) {
            return $point;
        }

        $frontAnchor = self::Project($plot, $chart, 0.0, 1.0, self::StackedFrontAxisGap);
        $currentLeft = self::Project($plot, $chart, 0.0, 1.0, $depth);
        $wallFrontLeft = self::Project($plot, $chart, 0.0, 1.0, 0.0);
        $wallBackLeft = self::Project($plot, $chart, 0.0, 1.0, 1.0);
        $wallBackRight = self::Project($plot, $chart, 1.0, 1.0, 1.0);
        $depthDx = $wallBackLeft->X - $wallFrontLeft->X;
        $backDx = $wallBackRight->X - $wallBackLeft->X;
        if (abs($depthDx) <= 0.000001 || abs($backDx) <= 0.000001) {
            return $point;
        }

        $depthSlope = ($wallBackLeft->Y - $wallFrontLeft->Y) / $depthDx;
        $currentLeftY = $frontAnchor->Y + $depthSlope * ($currentLeft->X - $frontAnchor->X);
        $backSlope = ($wallBackRight->Y - $wallBackLeft->Y) / $backDx;
        $point->Y = $currentLeftY + $backSlope * ($point->X - $currentLeft->X);
        return $point;
    }

    private static function DrawBandFace($canvas, $upper, $lower, $series, $fillColor, $usePattern)
    {
        $path = new SKPath();
        $path->MoveTo($upper[0]);
        for ($i = 1; $i < count($upper); $i++) {
            $path->LineTo($upper[$i]);
        }

        for ($i = count($lower) - 1; $i >= 0; $i--) {
            $path->LineTo($lower[$i]);
        }

        $path->Close();
        if ($usePattern && ChartPatternRenderer::DrawPath($canvas, $path, $series, $fillColor)) {
            $path->Dispose();
            return;
        }

        $fill = new SKPaint();
        $fill->Style = SKPaintStyle::Fill;
        $fill->Color = $fillColor;
        $fill->IsAntialias = true;
        $canvas->DrawPath($path, $fill);
        $fill->Dispose();
        $path->Dispose();
    }

    private static function DrawQuad($canvas, $first, $second, $third, $fourth, $paint)
    {
        $path = new SKPath();
        $path->MoveTo($first);
        $path->LineTo($second);
        $path->LineTo($third);
        $path->LineTo($fourth);
        $path->Close();
        $canvas->DrawPath($path, $paint);
        $path->Dispose();
    }

    private function DrawValueLabels($canvas, $area, $plot, $chart, $min, $max, $major)
    {
        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $normalized = ($value - $min) / ($max - $min);
            $axisX = $chart->ValueAxisOnRight ? 1.0 : 0.0;
            $axisDepth = $chart->ValueAxisOnRight ? 1.0 : 0.0;
            $point = self::Project($plot, $chart, $axisX, $normalized, $axisDepth);
            $label = ChartXmlParser::FormatLabel($value, $chart->ValueFormatCode, $this->dateSystem, $this->culture);
            $labelX = $point->X + ($chart->ValueAxisOnRight ? 7.0 : -7.0);
            $this->DrawText($canvas, $area, $label !== null ? rtrim($label) : '', $labelX, $point->Y, $chart->AxisLabelColor, !$chart->ValueAxisOnRight, false);
        }
    }

    private function DrawCategoryLabels($canvas, $area, $plot, $chart)
    {
        $monthlyTicks = Area3DDateAxisMath::BuildMonthlyTicks($chart, $this->dateSystem, $this->culture);
        if (count($monthlyTicks) > 0) {
            for ($i = 0; $i < count($monthlyTicks); $i++) {
                $point = self::Project($plot, $chart, $monthlyTicks[$i][0], 0.0, 0.0);
                $canvas->Save();
                $canvas->RotateDegrees((float)$chart->CategoryAxisTextRotationDeg, $point->X, $point->Y);
                $this->DrawText($canvas, $area, $monthlyTicks[$i][1], $point->X - 4.0, $point->Y, $chart->AxisLabelColor, true, false);
                $canvas->Restore();
            }

            return;
        }

        $count = max(count($chart->Categories), count($chart->CategoryValues));
        for ($i = 0; $i < $count; $i++) {
            $position = Area3DDateAxisMath::CategoryPosition($chart, $i, $count, $this->dateSystem);
            $point = self::Project($plot, $chart, $position, 0.0, 0.0);
            $label = $this->ResolveCategoryLabel($chart, $i);
            $this->DrawText($canvas, $area, $label, $point->X, $point->Y + 7.0, $chart->AxisLabelColor, false, true);
        }
    }

    private function DrawSeriesAxisLabels($canvas, $area, $plot, $chart)
    {
        if (!$chart->HasVisibleSeriesAxis) {
            return;
        }

        for ($i = count($chart->Series) - 1; $i >= 0; $i--) {
            $thickness = self::ResolveSeriesThickness(count($chart->Series));
            $depth = self::SeriesFrontDepth($i, count($chart->Series), $thickness);
            $point = self::Project($plot, $chart, 1.0, 0.0, $depth);
            $this->DrawText($canvas, $area, $chart->Series[$i]->Name, $point->X + 5.0, $point->Y, $chart->SeriesAxisLabelColor, false, false);
        }
    }

    private function ResolveCategoryLabel($chart, $index)
    {
        if ($index < count($chart->Categories) && !($chart->Categories[$index] === null || $chart->Categories[$index] === '')) {
            return $chart->Categories[$index];
        }

        if ($index < count($chart->CategoryValues) && $chart->CategoryValues[$index] !== null) {
            $formatCode = ($chart->CategoryFormatCode === null || $chart->CategoryFormatCode === '') ? 'mmm-yy' : $chart->CategoryFormatCode;
            $label = ChartXmlParser::FormatLabel($chart->CategoryValues[$index], $formatCode, $this->dateSystem, $this->culture);
            return $label !== null ? rtrim($label) : '';
        }

        return '';
    }

    private function DrawText($canvas, $clip, $text, $x, $y, $color, $alignRight, $alignTop)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = self::LabelSizePt;
        $fontContext = $this->context->GetFontContext($fontValue);
        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->IsAntialias = true;
        $paint->Typeface = $this->context->Fonts->Resolve($fontValue);
        $paint->TextSize = self::LabelSizePt;
        $paint->TextAlign = SKTextAlign::Left;
        $paint->SubpixelText = true;

        $drawX = $x;
        if ($alignRight) {
            $drawX -= $fontContext->Measure($text);
        } elseif ($alignTop) {
            $drawX -= $fontContext->Measure($text) * 0.5;
        }

        $metrics = $paint->FontMetrics;
        $baseline = $alignTop ? $y - $metrics->Ascent : $y - ($metrics->Ascent + $metrics->Descent) * 0.5;
        $canvas->Save();
        $canvas->ClipRect($clip);
        $runs = $fontContext->SplitRuns($text);
        $cursor = $drawX;
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $font = new SKFont($run->Typeface, self::LabelSizePt);
            $font->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $run->Text, $cursor, $baseline, $font, $color);
            $font->Dispose();
            $cursor += $run->WidthPt;
        }

        $canvas->Restore();
        $paint->Dispose();
    }

    private static function Project($plot, $chart, $x, $y, $depth)
    {
        $depthScale = self::Clamp($chart->DepthPercent / 100.0, 0.5, 2.0);
        $rotationScale = self::Clamp(abs($chart->RotationYDeg) / 20.0, 0.5, 1.5);
        $depthBlend = self::Clamp($depth * $depthScale * $rotationScale, 0.0, 1.0);
        $tiltScale = self::Clamp(abs($chart->RotationXDeg) / 15.0, 0.35, 1.5);
        $perspectiveScale = $chart->RightAngleAxes
            ? 0.85
            : self::Clamp($chart->PerspectiveDeg / 30.0, 0.5, 1.5);
        $pointCount = self::ResolvePointCount($chart);
        $compact = $chart->IsStacked || $chart->IsPercentStacked
            ? 0.0
            : self::Clamp((count($chart->Series) - $pointCount) / 3.0, 0.0, 1.0);

        $frontLeftX = $plot->Left + $plot->Width * self::Lerp(0.03, 0.23, $compact);
        $frontRightX = $plot->Left + $plot->Width * self::Lerp(0.86, 0.68, $compact);
        $backLeftX = $frontLeftX + $plot->Width * self::Lerp(0.19, 0.29, $compact) * (float)$perspectiveScale;
        $backRightX = $frontRightX + $plot->Width * self::Lerp(0.05, 0.13, $compact) * (float)$perspectiveScale;
        $frontBaseLeftY = $plot->Top + $plot->Height * self::Lerp(0.62, 0.85, $compact);
        $frontBaseRightY = $frontBaseLeftY + $plot->Height * self::Lerp(0.26, 0.16, $compact) * (float)$tiltScale;
        $backBaseLeftY = $frontBaseLeftY - $plot->Height * self::Lerp(0.13, 0.50, $compact) * (float)$tiltScale;
        $backBaseRightY = $backBaseLeftY + $plot->Height * self::Lerp(0.09, 0.10, $compact) * (float)$tiltScale;
        $frontHeight = $plot->Height * self::Lerp(0.56, 0.72, $compact);
        $backHeight = $plot->Height * self::Lerp(0.51, 0.42, $compact);

        $leftX = self::Lerp($frontLeftX, $backLeftX, $depthBlend);
        $rightX = self::Lerp($frontRightX, $backRightX, $depthBlend);
        $baseLeftY = self::Lerp($frontBaseLeftY, $backBaseLeftY, $depthBlend);
        $baseRightY = self::Lerp($frontBaseRightY, $backBaseRightY, $depthBlend);
        $height = self::Lerp($frontHeight, $backHeight, $depthBlend);
        $projectedX = self::Lerp($leftX, $rightX, $x);
        if ($chart->IsPercentStacked) {
            $perspectiveInset = $plot->Width * 0.02 * (float)(1.0 - self::Clamp($y, 0.0, 1.0));
            $projectedX += self::Lerp($perspectiveInset, -$perspectiveInset, (float)self::Clamp($x, 0.0, 1.0));
        }

        $baselineY = self::Lerp($baseLeftY, $baseRightY, $x);
        $rightHeight = $chart->IsPercentStacked ? $height * 1.10 : $height;
        $projectedHeight = self::Lerp($height, $rightHeight, (float)self::Clamp($x, 0.0, 1.0));
        return new SKPoint($projectedX, $baselineY - (float)$y * $projectedHeight);
    }

    private static function SeriesDepth($seriesIndex, $seriesCount)
    {
        if ($seriesCount <= 1) {
            return 0.5;
        }

        return (float)$seriesIndex / ($seriesCount - 1);
    }

    private static function ResolveSeriesThickness($seriesCount)
    {
        if ($seriesCount <= 0) {
            return 0.04;
        }

        return self::Clamp(0.24 / $seriesCount, 0.04, 0.075);
    }

    private static function ResolveStackedSeriesThickness($chart)
    {
        $depthPercent = $chart !== null ? $chart->DepthPercent : 100.0;
        if ($chart !== null && $chart->IsPercentStacked) {
            return self::Clamp(0.14 * $depthPercent / 100.0, 0.08, 0.22);
        }

        return self::Clamp(0.22 * $depthPercent / 100.0, 0.12, 0.32);
    }

    private static function SeriesFrontDepth($seriesIndex, $seriesCount, $thickness)
    {
        return self::FrontAxisGap + self::SeriesDepth($seriesIndex, $seriesCount) * (1.0 - self::FrontAxisGap - $thickness);
    }

    private static function ResolvePointCount($chart)
    {
        $count = max(count($chart->Categories), count($chart->CategoryValues));
        for ($i = 0; $i < count($chart->Series); $i++) {
            if (count($chart->Series[$i]->Values) > $count) {
                $count = count($chart->Series[$i]->Values);
            }
        }

        return $count;
    }

    private static function NormalizeValue($value, $min, $max)
    {
        return self::Clamp(($value - $min) / ($max - $min), 0.0, 1.0);
    }

    private static function ResolveMajorUnit($chart, $min, $max, $major)
    {
        if ($major <= 0.0) {
            return ($max - $min) / 4.0;
        }

        $pointCount = $chart !== null ? self::ResolvePointCount($chart) : 0;
        if ($chart !== null && count($chart->Series) > $pointCount) {
            return $major;
        }

        $steps = ($max - $min) / $major;
        if ($steps <= 5.0) {
            return $major;
        }

        return self::NiceNumber(($max - $min) / 4.0);
    }

    private static function NiceNumber($value)
    {
        if ($value <= 0.0) {
            return 1.0;
        }

        $exponent = floor(log10($value));
        $fraction = $value / pow(10.0, $exponent);
        if ($fraction < 1.5) {
            $nice = 1.0;
        } elseif ($fraction < 3.0) {
            $nice = 2.0;
        } elseif ($fraction < 7.0) {
            $nice = 5.0;
        } else {
            $nice = 10.0;
        }

        return $nice * pow(10.0, $exponent);
    }

    private static function Shade($color, $factor)
    {
        return new SKColor(
            (int)round($color->Red * $factor),
            (int)round($color->Green * $factor),
            (int)round($color->Blue * $factor),
            $color->Alpha);
    }

    private static function Lerp($start, $end, $amount)
    {
        return $start + ($end - $start) * (float)$amount;
    }

    private static function Clamp($value, $min, $max)
    {
        if ($value < $min) {
            return $min;
        }

        if ($value > $max) {
            return $max;
        }

        return $value;
    }
}
