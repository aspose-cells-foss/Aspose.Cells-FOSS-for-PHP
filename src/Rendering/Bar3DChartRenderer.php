<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSerialConverter;
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

final class Bar3DChartRenderer
{
    private const LabelSizePt = 9.0;

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
        if ($canvas === null || $chart === null || count($chart->Series) === 0 || $max <= $min) {
            return;
        }

        $denseGrid = $chart->HasMinorValueGridlines;
        $topLeftX = 0.086;
        $topLeftY = -0.018;
        $topRightX = 0.168;
        $topRightY = 0.076;
        $bottomLeftX = 0.104;
        $bottomLeftY = 0.512;
        $bottomRightX = 0.188;
        $bottomRightY = 0.743;
        if (self::UsesPerspectiveFan($chart)) {
            $topLeftX = 0.174;
            $topLeftY = 0.141;
            $topRightX = 0.169;
            $topRightY = -0.007;
            $bottomLeftX = 0.176;
            $bottomLeftY = 0.746;
            $bottomRightX = 0.177;
            $bottomRightY = 0.912;
        } elseif ($chart->HasDateCategoryAxis && !$chart->RightAngleAxes
            && $chart->PerspectiveDeg >= 100.0) {
            $topLeftX = 0.079;
            $topLeftY = 0.166;
            $topRightX = 0.231;
            $topRightY = 0.071;
            $bottomLeftX = 0.116;
            $bottomLeftY = 0.725;
            $bottomRightX = 0.284;
            $bottomRightY = 0.925;
        } elseif ($chart->HasDateCategoryAxis) {
            $topLeftX = 0.079;
            $topLeftY = 0.071;
            $topRightX = 0.133;
            $topRightY = 0.168;
            $bottomLeftX = 0.097;
            $bottomLeftY = 0.760;
            $bottomRightX = 0.153;
            $bottomRightY = 1.101;
        } elseif ($denseGrid) {
            $topLeftX = 0.073;
            $topLeftY = -0.077;
            $topRightX = 0.116;
            $topRightY = 0.023;
            $bottomLeftX = 0.094;
            $bottomLeftY = 0.746;
            $bottomRightX = 0.143;
            $bottomRightY = 1.002;
        } elseif ($chart->IsPercentStacked) {
            $topLeftX = 0.145;
            $topLeftY = -0.011;
            $topRightX = 0.117;
            $topRightY = 0.082;
            $bottomLeftX = 0.161;
            $bottomLeftY = 0.674;
            $bottomRightX = 0.143;
            $bottomRightY = 0.916;
        } elseif ($chart->IsStacked) {
            $topLeftX = 0.080;
            $topLeftY = -0.011;
            $topRightX = 0.143;
            $topRightY = 0.082;
            $bottomLeftX = 0.100;
            $bottomLeftY = 0.674;
            $bottomRightX = 0.165;
            $bottomRightY = 0.916;
        }

        $topLeft = new SKPoint(
            $plot->Left + $plot->Width * $topLeftX,
            $plot->Top + $plot->Height * $topLeftY);
        $topRight = new SKPoint(
            $plot->Right - $plot->Width * $topRightX,
            $plot->Top + $plot->Height * $topRightY);
        $bottomLeft = new SKPoint(
            $plot->Left + $plot->Width * $bottomLeftX,
            $plot->Top + $plot->Height * $bottomLeftY);
        $bottomRight = new SKPoint(
            $plot->Right - $plot->Width * $bottomRightX,
            $plot->Top + $plot->Height * $bottomRightY);
        self::ResolveSideWall($plot, $topLeft, $bottomLeft, $wallFrontTopLeft, $wallFrontBottomLeft);
        $sideWallImage = $chart->HasDateCategoryAxis ? null : $chart->SideWallImageData;
        $wallFallback = $chart->HasDateCategoryAxis ? SKColors::Transparent() : SKColors::White();
        Area3DWallFillRenderer::Draw(
            $canvas,
            $wallFrontBottomLeft,
            $wallFrontTopLeft,
            $bottomLeft,
            $topLeft,
            $bottomRight,
            $topRight,
            $sideWallImage,
            $chart->BackWallImageData,
            $wallFallback);
        if ($chart->HasDateCategoryAxis) {
            $floorOffset = new SKPoint(
                $wallFrontBottomLeft->X - $bottomLeft->X,
                $wallFrontBottomLeft->Y - $bottomLeft->Y);
            $floorFrontRight = self::Offset($bottomRight, $floorOffset);
            Area3DWallFillRenderer::DrawFloor(
                $canvas,
                $bottomLeft,
                $bottomRight,
                $floorFrontRight,
                $wallFrontBottomLeft,
                $chart->BackWallImageData,
                new SKColor(0xE5, 0xD4, 0xA2));
        }

        self::DrawSideWall($canvas, $wallFrontTopLeft, $topLeft, $bottomLeft, $wallFrontBottomLeft);
        $renderedMajor = $denseGrid || $chart->HasDateCategoryAxis
            ? $major
            : self::NiceNumber(($max - $min) / 4.0);
        $this->DrawGridAndValueAxis($canvas, $area, $topLeft, $topRight, $bottomLeft, $bottomRight,
            $wallFrontTopLeft, $wallFrontBottomLeft, $chart, $min, $max, $renderedMajor);
        $this->DrawBars($canvas, $topLeft, $topRight, $bottomLeft, $bottomRight,
            $wallFrontTopLeft, $wallFrontBottomLeft, $chart, $min, $max);
        $this->DrawCategoryLabels($canvas, $area, $wallFrontTopLeft, $wallFrontBottomLeft, $chart);
    }

    private static function UsesPerspectiveFan($chart)
    {
        return $chart->IsPercentStacked
            && !$chart->RightAngleAxes
            && abs($chart->RotationXDeg) < 0.01
            && $chart->PerspectiveDeg > 0.0;
    }

    private function DrawGridAndValueAxis($canvas, $area, $topLeft, $topRight, $bottomLeft, $bottomRight, $wallFrontTopLeft, $wallFrontBottomLeft, $chart, $min, $max, $major)
    {
        self::DrawMinorGridlines($canvas, $topLeft, $topRight, $bottomLeft, $bottomRight, $chart, $min, $max, $major);
        $grid = new SKPaint();
        $grid->Style = SKPaintStyle::Stroke;
        $grid->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $grid->StrokeWidth = 0.75;
        $grid->IsAntialias = true;

        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $ratio = ($value - $min) / ($max - $min);
            $projectedRatio = self::ProjectValueRatio((float)$ratio);
            $top = self::Interpolate($topLeft, $topRight, $projectedRatio);
            $bottom = self::Interpolate($bottomLeft, $bottomRight, $projectedRatio);
            $canvas->DrawLine($top, $bottom, $grid);
            if ($chart->HasDateCategoryAxis && $chart->PerspectiveDeg >= 100.0 && $i % 2 != 0) {
                continue;
            }

            $label = ChartXmlParser::FormatLabel(
                $value / $chart->ValueDisplayUnitDivisor,
                $chart->ValueFormatCode,
                $this->dateSystem,
                $this->culture);
            $this->DrawCenteredText($canvas, $area, $label !== null ? rtrim($label) : '',
                $bottom->X, $bottom->Y + 11.0, $chart->AxisLabelColor);
        }

        $categoryTicks = $chart->HasMinorValueGridlines ? 6 : 3;
        for ($i = 0; $i <= $categoryTicks; $i++) {
            $tick = self::Interpolate($wallFrontTopLeft, $wallFrontBottomLeft, $i / (float)$categoryTicks);
            $canvas->DrawLine($tick->X, $tick->Y, $tick->X - 2.8, $tick->Y, $grid);
        }

        $grid->Dispose();

        self::DrawValueAxisLine($canvas, $bottomLeft, $bottomRight, $chart, $min, $max, $major);

        if (!($chart->ValueDisplayUnitLabel === null || $chart->ValueDisplayUnitLabel === '')) {
            $this->DrawText($canvas, $area, $chart->ValueDisplayUnitLabel,
                $bottomRight->X - 5.0, $bottomRight->Y + 27.0, $chart->AxisLabelColor, true);
        }
    }

    private static function DrawSideWall($canvas, $frontTopLeft, $backTopLeft, $backBottomLeft, $frontBottomLeft)
    {
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $paint->StrokeWidth = 0.75;
        $paint->IsAntialias = true;
        $canvas->DrawLine($frontTopLeft, $backTopLeft, $paint);
        $canvas->DrawLine($frontTopLeft, $frontBottomLeft, $paint);
        $canvas->DrawLine($frontBottomLeft, $backBottomLeft, $paint);
        $paint->Dispose();
    }

    private static function ResolveSideWall($plot, $backTopLeft, $backBottomLeft, &$frontTopLeft, &$frontBottomLeft)
    {
        $frontTopLeft = new SKPoint(
            $backTopLeft->X - max(4.0, $plot->Width * 0.022),
            $backTopLeft->Y + $plot->Height * 0.010);
        $frontBottomLeft = new SKPoint(
            $backBottomLeft->X - max(3.0, $plot->Width * 0.014),
            $backBottomLeft->Y + $plot->Height * 0.025);
    }

    private static function DrawMinorGridlines($canvas, $topLeft, $topRight, $bottomLeft, $bottomRight, $chart, $min, $max, $major)
    {
        if (!$chart->HasMinorValueGridlines || $major <= 0.0) {
            return;
        }

        $minor = $chart->MinorUnit !== null ? $chart->MinorUnit : $major / 5.0;
        if ($minor <= 0.0) {
            return;
        }

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->MinorValueGridlineColor;
        $paint->StrokeWidth = (float)$chart->MinorValueGridlineWidthPt;
        $paint->IsAntialias = true;
        $steps = (int)round(($max - $min) / $minor);
        for ($i = 1; $i < $steps; $i++) {
            $value = $min + $i * $minor;
            $majorRatio = ($value - $min) / $major;
            if (abs($majorRatio - round($majorRatio)) < 0.0001) {
                continue;
            }

            $ratio = self::ProjectValueRatio((float)(($value - $min) / ($max - $min)));
            $canvas->DrawLine(
                self::Interpolate($topLeft, $topRight, $ratio),
                self::Interpolate($bottomLeft, $bottomRight, $ratio),
                $paint);
        }

        $paint->Dispose();
    }

    private static function DrawValueAxisLine($canvas, $bottomLeft, $bottomRight, $chart, $min, $max, $major)
    {
        if (!$chart->HasValueAxisLine) {
            return;
        }

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->ValueAxisLineColor;
        $paint->StrokeWidth = (float)$chart->ValueAxisLineWidthPt;
        $paint->IsAntialias = true;

        $canvas->DrawLine($bottomLeft, $bottomRight, $paint);
        if ($chart->ValueAxisMajorTickMark === 'none') {
            $paint->Dispose();
            return;
        }

        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $ratio = self::ProjectValueRatio($i / (float)$steps);
            $tick = self::Interpolate($bottomLeft, $bottomRight, $ratio);
            $startY = $tick->Y;
            $endY = $tick->Y + 5.6;
            if ($chart->ValueAxisMajorTickMark === 'cross') {
                $startY -= 2.8;
                $endY -= 2.8;
            } elseif ($chart->ValueAxisMajorTickMark === 'in') {
                $endY = $tick->Y - 5.6;
            }

            $canvas->DrawLine($tick->X, $startY, $tick->X, $endY, $paint);
        }

        $paint->Dispose();
    }

    private function DrawBars($canvas, $topLeft, $topRight, $bottomLeft, $bottomRight, $wallFrontTopLeft, $wallFrontBottomLeft, $chart, $min, $max)
    {
        $pointCount = count($chart->Categories);
        if ($pointCount <= 0) {
            return;
        }

        $leftHeight = self::Distance($topLeft, $bottomLeft);
        $categorySlotCount = $this->ResolveCategorySlotCount($chart, $pointCount);
        $slot = $leftHeight / $categorySlotCount;
        $thicknessDirection = self::ResolveBarThicknessDirection($topLeft, $bottomLeft);
        if ($chart->IsStacked || $chart->IsPercentStacked) {
            $this->DrawStackedBars($canvas, $topLeft, $topRight, $bottomLeft, $bottomRight,
                $wallFrontTopLeft, $wallFrontBottomLeft,
                $thicknessDirection, $chart, $min, $max, $slot);
            return;
        }

        $barWidth = self::ResolveClusteredBarWidth($chart, $slot);
        for ($i = 0; $i < $pointCount; $i++) {
            $categoryIndex = $pointCount - 1 - $i;
            for ($s = count($chart->Series) - 1; $s >= 0; $s--) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $ratio = ($series->Values[$i] - $min) / ($max - $min);
                $seriesSpacing = self::ResolveClusteredSeriesSpacing($chart, $slot, $barWidth);
                $seriesOffset = (count($chart->Series) - 1 - $s - 1.0) * $seriesSpacing;
                $categoryRatio = $this->ResolveClusteredCategoryRatio(
                    $chart, $i, $categoryIndex, $pointCount, $categorySlotCount,
                    $seriesSpacing, $seriesOffset);
                $left = self::Interpolate($topLeft, $bottomLeft, $categoryRatio);
                $right = self::Interpolate($topRight, $bottomRight, $categoryRatio);
                self::ShiftBarToWallMidpoint(
                    $left,
                    self::Interpolate($wallFrontTopLeft, $wallFrontBottomLeft, $categoryRatio),
                    $left,
                    $right);
                $end = self::Interpolate($left, $right, self::ProjectValueRatio((float)$ratio));
                self::DrawBarSegment($canvas, $left, $end, $thicknessDirection,
                    $series->Color, $barWidth, true);
            }
        }
    }

    private function DrawStackedBars($canvas, $topLeft, $topRight, $bottomLeft, $bottomRight, $wallFrontTopLeft, $wallFrontBottomLeft, $thicknessDirection, $chart, $min, $max, $slot)
    {
        $pointCount = count($chart->Categories);
        $widthFactor = self::UsesPerspectiveFan($chart) ? 0.42 : 0.5;
        $barWidth = max(3.0, $slot * $widthFactor);
        for ($i = 0; $i < $pointCount; $i++) {
            $categoryIndex = $pointCount - 1 - $i;
            $categoryRatio = $this->ResolveCategoryRatio($chart, $i, $categoryIndex, $pointCount);
            $left = self::Interpolate($topLeft, $bottomLeft, $categoryRatio);
            $right = self::Interpolate($topRight, $bottomRight, $categoryRatio);
            self::ShiftBarToWallMidpoint(
                $left,
                self::Interpolate($wallFrontTopLeft, $wallFrontBottomLeft, $categoryRatio),
                $left,
                $right);
            $positive = 0.0;
            $negative = 0.0;
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $value = self::NormalizeStackedValue($chart, $i, $series->Values[$i]);
                $startValue = $value >= 0.0 ? $positive : $negative;
                $endValue = $startValue + $value;
                if ($value >= 0.0) {
                    $positive = $endValue;
                } else {
                    $negative = $endValue;
                }

                $startRatio = self::ProjectValueRatio((float)(($startValue - $min) / ($max - $min)));
                $endRatio = self::ProjectValueRatio((float)(($endValue - $min) / ($max - $min)));
                self::DrawBarSegment(
                    $canvas,
                    self::Interpolate($left, $right, $startRatio),
                    self::Interpolate($left, $right, $endRatio),
                    $thicknessDirection,
                    $series->Color,
                    $barWidth,
                    self::IsLastStackedSegment($chart, $i, $s, $value));
            }
        }
    }

    private static function NormalizeStackedValue($chart, $categoryIndex, $value)
    {
        if (!$chart->IsPercentStacked) {
            return $value;
        }

        $total = 0.0;
        $positive = $value >= 0.0;
        for ($i = 0; $i < count($chart->Series); $i++) {
            $series = $chart->Series[$i];
            if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                continue;
            }

            $candidate = $series->Values[$categoryIndex];
            if (($candidate >= 0.0) != $positive) {
                continue;
            }

            $total += abs($candidate);
        }

        if ($total <= 0.0) {
            return 0.0;
        }

        return $value / $total;
    }

    private static function ShiftBarToWallMidpoint($wallBack, $wallFront, &$left, &$right)
    {
        $offset = new SKPoint(
            ($wallFront->X - $wallBack->X) * 0.5,
            ($wallFront->Y - $wallBack->Y) * 0.5);
        $left = self::Offset($left, $offset);
        $right = self::Offset($right, $offset);
    }

    private static function IsLastStackedSegment($chart, $categoryIndex, $seriesIndex, $value)
    {
        $positive = $value >= 0.0;
        for ($i = $seriesIndex + 1; $i < count($chart->Series); $i++) {
            $series = $chart->Series[$i];
            if ($categoryIndex >= count($series->Values) || $series->Values[$categoryIndex] === null) {
                continue;
            }

            $laterValue = $series->Values[$categoryIndex];
            if (($laterValue >= 0.0) == $positive && $laterValue != 0.0) {
                return false;
            }
        }

        return true;
    }

    private function ResolveCategorySlotCount($chart, $pointCount)
    {
        if (!$chart->HasDateCategoryAxis) {
            return $pointCount;
        }

        if (!ChartDateAxisMath::TryGetDateAxisRange($chart, $minDate, $maxDate)) {
            return $pointCount;
        }

        $min = DateSerialConverter::FromSerial($minDate, $this->dateSystem);
        $max = DateSerialConverter::FromSerial($maxDate, $this->dateSystem);
        $monthCount = ((int)$max->format('Y') - (int)$min->format('Y')) * 12 + (int)$max->format('n') - (int)$min->format('n') + 1;
        return max($pointCount, $monthCount);
    }

    private function ResolveCategoryRatio($chart, $sourceIndex, $reversedIndex, $pointCount)
    {
        if ($chart->HasDateCategoryAxis
            && $sourceIndex < count($chart->CategoryValues)
            && $chart->CategoryValues[$sourceIndex] !== null) {
            if (ChartDateAxisMath::TryGetDateAxisPlotRange(
                $chart, $this->dateSystem, $plotMinDate, $plotMaxDate)) {
                $serial = ChartDateAxisMath::NormalizePointSerial(
                    $chart, $chart->CategoryValues[$sourceIndex], $this->dateSystem);
                $ratio = ($serial - $plotMinDate) / ($plotMaxDate - $plotMinDate);
                $ratio = max(0.0, min(1.0, $ratio));
                return $chart->CategoryAxisMinMax ? 1.0 - (float)$ratio : (float)$ratio;
            }
        }

        return ($reversedIndex + 0.5) / $pointCount;
    }

    private function ResolveClusteredCategoryRatio($chart, $sourceIndex, $reversedIndex, $pointCount, $categorySlotCount, $seriesSpacing, $seriesOffset)
    {
        $baseRatio = $this->ResolveCategoryRatio($chart, $sourceIndex, $reversedIndex, $pointCount);
        if ($chart->HasDateCategoryAxis) {
            $minimumBaseRatio = (0.6 + $seriesSpacing) / $categorySlotCount;
            if ($baseRatio < $minimumBaseRatio) {
                $baseRatio = $minimumBaseRatio;
            }
        }

        return $baseRatio + $seriesOffset / $categorySlotCount;
    }

    private static function ResolveClusteredSeriesSpacing($chart, $slot, $barWidth)
    {
        if (!$chart->HasDateCategoryAxis) {
            return 0.20;
        }

        return $barWidth / max(0.01, $slot);
    }

    private static function ResolveClusteredBarWidth($chart, $slot)
    {
        if ($chart->HasDateCategoryAxis) {
            return max(1.0, $slot * 0.25);
        }

        return max(2.5, $slot * 0.25);
    }

    private static function DrawBarSegment($canvas, $start, $end, $thicknessDirection, $color, $width, $drawEndFace)
    {
        $dx = $end->X - $start->X;
        $dy = $end->Y - $start->Y;
        $length = (float)sqrt($dx * $dx + $dy * $dy);
        if ($length <= 0.0) {
            return;
        }

        $halfWidth = $width * 0.5;
        $halfThickness = new SKPoint(
            $thicknessDirection->X * $halfWidth,
            $thicknessDirection->Y * $halfWidth);
        $frontTopStart = new SKPoint(
            $start->X - $halfThickness->X,
            $start->Y - $halfThickness->Y);
        $frontTopEnd = new SKPoint(
            $end->X - $halfThickness->X,
            $end->Y - $halfThickness->Y);
        $frontBottomStart = self::Offset($start, $halfThickness);
        $frontBottomEnd = self::Offset($end, $halfThickness);
        $depth = min(2.0, max(0.75, $width * 0.20));
        $depthOffset = new SKPoint($depth, -$depth * 0.55);
        $backTopStart = self::Offset($frontTopStart, $depthOffset);
        $backTopEnd = self::Offset($frontTopEnd, $depthOffset);

        self::DrawBarFace($canvas, self::Shade($color, 0.76),
            $frontTopStart, $frontTopEnd, $backTopEnd, $backTopStart);
        self::DrawBarFace($canvas, $color,
            $frontTopStart, $frontTopEnd, $frontBottomEnd, $frontBottomStart);
        if ($drawEndFace) {
            $backBottomEnd = self::Offset($frontBottomEnd, $depthOffset);
            self::DrawBarFace($canvas, self::Shade($color, 0.64),
                $frontTopEnd, $backTopEnd, $backBottomEnd, $frontBottomEnd);
        }
    }

    private static function DrawBarFace($canvas, $color, $first, $second, $third, $fourth)
    {
        $path = new SKPath();
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Fill;
        $paint->Color = $color;
        $paint->IsAntialias = true;
        $path->MoveTo($first);
        $path->LineTo($second);
        $path->LineTo($third);
        $path->LineTo($fourth);
        $path->Close();
        $canvas->DrawPath($path, $paint);
        $paint->Dispose();
        $path->Dispose();
    }

    private static function Offset($point, $offset)
    {
        return new SKPoint($point->X + $offset->X, $point->Y + $offset->Y);
    }

    private static function Shade($color, $factor)
    {
        return new SKColor(
            (int)round($color->Red * $factor),
            (int)round($color->Green * $factor),
            (int)round($color->Blue * $factor),
            $color->Alpha);
    }

    private function DrawCategoryLabels($canvas, $area, $topLeft, $bottomLeft, $chart)
    {
        if ($chart->HasDateCategoryAxis && $this->DrawDateCategoryLabels($canvas, $area, $topLeft, $bottomLeft, $chart)) {
            return;
        }

        $count = count($chart->Categories);
        $labelStep = $chart->HasMinorValueGridlines ? 1 : 2;
        for ($i = 0; $i < $count; $i += $labelStep) {
            $categoryIndex = $count - 1 - $i;
            $point = self::Interpolate($topLeft, $bottomLeft, ($categoryIndex + 0.5) / $count);
            $this->DrawText($canvas, $area, $chart->Categories[$i], $point->X - 8.0, $point->Y, $chart->AxisLabelColor, true);
        }
    }

    private function DrawDateCategoryLabels($canvas, $area, $topLeft, $bottomLeft, $chart)
    {
        if (!ChartDateAxisMath::TryGetDateAxisRange($chart, $minDate, $maxDate)
            || !ChartDateAxisMath::TryGetDateAxisPlotRange(
                $chart, $this->dateSystem, $plotMinDate, $plotMaxDate)) {
            return false;
        }

        $slotCount = $this->ResolveCategorySlotCount($chart, count($chart->Categories));
        $monthStep = $chart->PerspectiveDeg >= 100.0 ? 4 : ($slotCount > 8 ? 3 : 1);
        $current = ChartDateAxisMath::StartOfMonth($minDate, $this->dateSystem);
        $last = ChartDateAxisMath::StartOfMonth($maxDate, $this->dateSystem);
        while ($current <= $last + 0.5) {
            $ratio = ($current - $plotMinDate) / ($plotMaxDate - $plotMinDate);
            $ratio = max(0.0, min(1.0, $ratio));
            $categoryRatio = $chart->CategoryAxisMinMax ? 1.0 - (float)$ratio : (float)$ratio;
            $point = self::Interpolate($topLeft, $bottomLeft, $categoryRatio);
            $formatCode = ($chart->CategoryFormatCode === null || $chart->CategoryFormatCode === '')
                ? 'mmm-yy'
                : $chart->CategoryFormatCode;
            $label = ChartXmlParser::FormatLabel(
                $current, $formatCode, $this->dateSystem, $this->culture);
            $this->DrawText($canvas, $area, $label !== null ? rtrim($label) : '',
                $point->X - 8.0, $point->Y, $chart->AxisLabelColor, true);
            $current = ChartDateAxisMath::AddMonthsSerial($current, $monthStep, $this->dateSystem);
        }

        return true;
    }

    private static function Interpolate($start, $end, $ratio)
    {
        return new SKPoint(
            $start->X + ($end->X - $start->X) * $ratio,
            $start->Y + ($end->Y - $start->Y) * $ratio);
    }

    private static function ProjectValueRatio($ratio)
    {
        return $ratio / (1.0 + 0.20 * (1.0 - $ratio));
    }

    private static function Distance($first, $second)
    {
        $dx = $second->X - $first->X;
        $dy = $second->Y - $first->Y;
        return (float)sqrt($dx * $dx + $dy * $dy);
    }

    private static function ResolveBarThicknessDirection($wallTopLeft, $wallBottomLeft)
    {
        $length = self::Distance($wallTopLeft, $wallBottomLeft);
        return new SKPoint(
            ($wallBottomLeft->X - $wallTopLeft->X) / $length,
            ($wallBottomLeft->Y - $wallTopLeft->Y) / $length);
    }

    private function DrawText($canvas, $clip, $text, $x, $y, $color, $alignRight)
    {
        if ($text === null || $text === '') return;
        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = self::LabelSizePt;
        $fontContext = $this->context->GetFontContext($fontValue);
        $drawX = $alignRight ? $x - $fontContext->Measure($text) : $x;
        $runs = $fontContext->SplitRuns($text);
        $canvas->Save();
        $canvas->ClipRect($clip);
        for ($i = 0; $i < count($runs); $i++) {
            $font = new SKFont($runs[$i]->Typeface, self::LabelSizePt);
            PdfTextPathRenderer::DrawText($canvas, $runs[$i]->Text, $drawX, $y + self::LabelSizePt * 0.35, $font, $color);
            $font->Dispose();
            $drawX += $runs[$i]->WidthPt;
        }

        $canvas->Restore();
    }

    private function DrawCenteredText($canvas, $clip, $text, $x, $y, $color)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = self::LabelSizePt;
        $fontContext = $this->context->GetFontContext($fontValue);
        $this->DrawText($canvas, $clip, $text, $x + $fontContext->Measure($text) * 0.5, $y, $color, true);
    }

    private static function NiceNumber($value)
    {
        if ($value <= 0.0) return 1.0;
        $exponent = floor(log10($value));
        $fraction = $value / pow(10.0, $exponent);
        $nice = $fraction < 1.5 ? 1.0 : ($fraction < 3.0 ? 2.0 : ($fraction < 7.0 ? 5.0 : 10.0));
        return $nice * pow(10.0, $exponent);
    }
}
