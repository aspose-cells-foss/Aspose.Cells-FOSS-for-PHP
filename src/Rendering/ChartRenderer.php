<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\Skia\SKBlurStyle;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKMaskFilter;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;
use Aspose\Cells_FOSS\Skia\SKStrokeCap;
use Aspose\Cells_FOSS\Skia\SKStrokeJoin;
use Aspose\Cells_FOSS\Skia\SKTextAlign;

final class ChartRenderer
{
    const StandardDateAxisGridlineRightExtensionPt = 8.0;

    const StackedDateAxisGridlineRightExtensionPt = 10.0;

    private static $gridColor = null;

    private static $borderColor = null;

    private $LabelSizePt = 9.0;

    private $context;

    private $dateSystem;

    private $chartCulture;

    private $suppressCategoryAxisLabels = false;

    private static function GridColor()
    {
        if (self::$gridColor === null) {
            self::$gridColor = new SKColor(0xD9, 0xD9, 0xD9);
        }

        return self::$gridColor;
    }

    private static function BorderColor()
    {
        if (self::$borderColor === null) {
            self::$borderColor = new SKColor(0xD9, 0xD9, 0xD9);
        }

        return self::$borderColor;
    }

    public function __construct($context)
    {
        $this->context = $context;
        $this->dateSystem = $context->Workbook !== null && $context->Workbook->Settings !== null
            ? $context->Workbook->Settings->DateSystem
            : DateSystem::Windows1900;
    }

    public function Draw($canvas, $page, $area, $chart, $suppressCategoryAxisLabels = false)
    {
        $this->chartCulture = $chart !== null && $chart->Culture !== null ? $chart->Culture : $this->context->Culture;
        $this->LabelSizePt = (float)($chart->TextSizePt ?? 9.0);
        $this->suppressCategoryAxisLabels = $suppressCategoryAxisLabels;

        $this->DrawFrame($canvas, $area, $chart);

        if ($page !== null && $chart->IsArea3D && ($area->Left < 0.0 || $area->Top < 0.0)) {
            return;
        }

        if ($chart->Kind === ChartKind::Pie) {
            $this->DrawPie($canvas, $page, $area, $chart);
            return;
        }

        if (count($chart->Series) === 0) {
            $this->DrawTitle($canvas, $page, $area, $chart, $area->MidY);
            return;
        }

        $plot = $this->ResolvePlotRect($area, $chart);
        $seriesPlot = $plot;

        if ($chart->Kind === ChartKind::Scatter && $chart->ScatterVaryColors
            && $chart->ScatterCategoryAxisDeleted && $chart->ScatterValueAxisDeleted
            && $chart->ScatterTitleOverlay && $chart->RoundedCorners) {
            $this->DrawHiddenAxisScatter($canvas, $page, $area, $chart);
            return;
        }

        $min = 0.0;
        $max = 0.0;
        $major = 0.0;
        self::ComputeValueAxis($chart, $min, $max, $major);

        $titleCenterY = $area->Top + ($seriesPlot->Top - $area->Top) * 0.5;
        if ($chart->Kind === ChartKind::Scatter && !$chart->HasManualTitle) {
            $titleCenterY = $area->Top + ($seriesPlot->Top - $area->Top)
                * ($chart->LegendPosition === null ? 0.45 : 0.38);
        }
        if ($chart->HasDataTable) {
            $titleCenterY = $area->Top + $area->Height * 0.08;
        } elseif ($chart->Kind === ChartKind::Column && $chart->IsBar3D) {
            $titleCenterY = $area->Top + $area->Height * 0.08;
        } elseif ($chart->Kind === ChartKind::Column && $chart->IsPercentStacked && !$chart->IsBar3D) {
            $titleCenterY = $area->Top + $area->Height * 0.06;
        }
        if ($chart->Kind === ChartKind::Radar) {
            $titleCenterY = $area->Top + ($seriesPlot->Top - $area->Top) * 0.55;
        } elseif ($chart->Kind === ChartKind::Area && $chart->IsArea3D && $chart->IsStacked
            && $chart->LegendPosition === 't') {
            $titleCenterY = $area->Top + $area->Height * 0.09;
        }

        $this->DrawTitle($canvas, $page, $area, $chart, LineChartDecorationLayout::TitleCenterY($area, $chart, $titleCenterY));
        if ($chart->Kind === ChartKind::Scatter) {
            $this->DrawValueAxis($canvas, $page, $area, $seriesPlot, $chart, $min, $max, $major);
            $this->DrawScatterCategoryAxis($canvas, $page, $seriesPlot, $chart, $min, $max);
            if ($chart->IsBubble) {
                BubbleChartMath::DrawSeries($canvas, $seriesPlot, $chart, $min, $max);
            } else {
                $this->DrawScatterSeries($canvas, $seriesPlot, $chart, $min, $max);
            }
            $this->DrawPlotAreaBorder($canvas, $page, $area, $seriesPlot, $chart);
            $this->DrawAxisTitles($canvas, $page, $area, $seriesPlot, $chart);
            if ($chart->ValueDisplayUnitLabel !== null && $chart->ValueDisplayUnitLabel !== '') {
                $unitX = $seriesPlot->Left - 44.0;
                $unitY = $seriesPlot->Top + $seriesPlot->Height * 0.2;
                if ($chart->HasManualValueDisplayUnitLabel) {
                    $unitX = $area->Left + (float)($chart->ValueDisplayUnitLabelX * $area->Width) + 8.0;
                    $unitY = $area->Top + (float)($chart->ValueDisplayUnitLabelY * $area->Height)
                        + $this->MeasureText($chart->ValueDisplayUnitLabel, 10.0) * 0.5 + 3.5;
                }
                $this->DrawRotatedText($canvas, $page, $area, $chart->ValueDisplayUnitLabel,
                    $unitX, $unitY, 10.0, false,
                    $chart->AxisLabelColor, TextAlign::Center, VAlign::Middle, -90.0);
            }
            $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
            return;
        }

        if ($chart->Kind === ChartKind::Radar) {
            $this->DrawRadar($canvas, $page, $area, $seriesPlot, $chart, $min, $max, $major);
            $this->DrawPlotAreaBorder($canvas, $page, $area, $seriesPlot, $chart);
            $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
            return;
        }

        if ($chart->Kind === ChartKind::Bar) {
            if ($chart->IsBar3D) {
                $bar3DRenderer = new Bar3DChartRenderer($this->context, $this->dateSystem, $this->chartCulture);
                $bar3DRenderer->Draw($canvas, $area, $seriesPlot, $chart, $min, $max, $major);
                $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
                return;
            }

            $this->DrawBarValueAxis($canvas, $page, $area, $seriesPlot, $chart, $min, $max, $major);
            $this->DrawSeries($canvas, $page, $seriesPlot, $chart, $min, $max);
            $this->DrawBarCategoryAxis($canvas, $page, $area, $seriesPlot, $chart);
            $this->DrawPlotAreaBorder($canvas, $page, $area, $seriesPlot, $chart);
            $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
            return;
        } elseif ($chart->Kind === ChartKind::Column && $chart->IsBar3D) {
            $column3DRenderer = new Column3DChartRenderer($this->context, $this->dateSystem, $this->chartCulture);
            $column3DRenderer->Draw($canvas, $area, $chart, $min, $max, $major);
            $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
            return;
        } elseif ($chart->Kind === ChartKind::Area && $chart->IsArea3D) {
            $area3DRenderer = new Area3DChartRenderer($this->context, $this->dateSystem, $this->chartCulture);
            $area3DRenderer->Draw($canvas, $area, $seriesPlot, $chart, $min, $max, $major);
            $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
            return;
        } elseif ($chart->Kind === ChartKind::Line && $chart->IsLine3D) {
            $line3DRenderer = new Line3DChartRenderer($this->context, $this->dateSystem, $this->chartCulture);
            $line3DRenderer->Draw($canvas, $page, $area, $seriesPlot, $chart, $min, $max, $major);
            $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
            return;
        } else {
            $this->DrawValueAxis($canvas, $page, $area, $seriesPlot, $chart, $min, $max, $major);
            $this->DrawCategoryAxis($canvas, $page, $area, $seriesPlot, $chart);
            $this->DrawAxisTitles($canvas, $page, $area, $seriesPlot, $chart);
        }
        $this->DrawSeries($canvas, $page, $seriesPlot, $chart, $min, $max);
        $this->DrawPlotAreaBorder($canvas, $page, $area, $seriesPlot, $chart);
        $this->DrawDataTable($canvas, $page, $area, $seriesPlot, $chart);
        $this->DrawLegend($canvas, $page, $area, $seriesPlot, $chart);
    }

    public function DrawFrame($canvas, $area, $chart)
    {
        if ($canvas === null) {
            return;
        }

        if ($chart !== null && $chart->HasChartAreaGlow) {
            $sigma = (float)($chart->ChartAreaGlowRadiusPt / 4.0);
            $spread = (float)($chart->ChartAreaGlowRadiusPt / 3.0);
            $source = $chart->ChartAreaGlowColor;
            $opacity = (int)min(255, (int)round($source->Alpha * 1.25));
            $glow = new SKPaint();
            $glow->Style = SKPaintStyle::Fill;
            $glow->Color = new SKColor($source->Red, $source->Green, $source->Blue, $opacity);
            $glow->IsAntialias = true;
            $blur = SKMaskFilter::CreateBlur(SKBlurStyle::Normal, $sigma);
            $glow->MaskFilter = $blur;
            $glowRect = new SKRect($area->Left - $spread, $area->Top - $spread,
                $area->Left - $spread + $area->Width + $spread * 2.0,
                $area->Top - $spread + $area->Height + $spread * 2.0);
            if ($chart->ChartAreaSoftEdgeRadiusPt > 0.0) {
                $cornerRadius = (float)min($chart->ChartAreaSoftEdgeRadiusPt,
                    min($glowRect->Width, $glowRect->Height) * 0.5);
                $canvas->DrawRoundRect($glowRect, $cornerRadius, $cornerRadius, $glow);
            } else {
                $canvas->DrawRect($glowRect, $glow);
            }

            $glow->Dispose();
            $blur->Dispose();
        }

        if ($chart !== null && $chart->HasChartAreaGradient
            && count($chart->ChartAreaGradientColors) >= 2
            && count($chart->ChartAreaGradientColors) === count($chart->ChartAreaGradientPositions)) {
            $radians = $chart->ChartAreaGradientAngleDeg * M_PI / 180.0;
            $xOffset = (float)(cos($radians) * $area->Width * 0.5);
            $yOffset = (float)(sin($radians) * $area->Height * 0.5);
            $start = new SKPoint($area->MidX - $xOffset, $area->MidY - $yOffset);
            $end = new SKPoint($area->MidX + $xOffset, $area->MidY + $yOffset);
            $shader = SKShader::CreateLinearGradient(
                $start,
                $end,
                $chart->ChartAreaGradientColors,
                $chart->ChartAreaGradientPositions,
                SKShaderTileMode::Clamp);
            $gradientPaint = new SKPaint();
            $gradientPaint->Style = SKPaintStyle::Fill;
            $gradientPaint->IsAntialias = false;
            $gradientPaint->Shader = $shader;
            $canvas->DrawRect($area, $gradientPaint);
            $gradientPaint->Dispose();
            $shader->Dispose();
        } else {
            if (!ChartPatternRenderer::DrawChartArea($canvas, $area, $chart)) {
                $background = SKColors::White();
                if ($chart !== null && $chart->HasChartAreaFill) {
                    $background = $chart->ChartAreaFillColor;
                }

                if ($chart !== null && $chart->ChartAreaSoftEdgeRadiusPt > 0.0) {
                    ChartSoftEdgeRenderer::DrawSolidFill($canvas, $area, $background, $chart->ChartAreaSoftEdgeRadiusPt);
                } else {
                    $paint = new SKPaint();
                    $paint->Style = SKPaintStyle::Fill;
                    $paint->Color = $background;
                    $paint->IsAntialias = false;
                    $canvas->DrawRect($area, $paint);
                    $paint->Dispose();
                }
            }
        }

        $borderColor = $chart !== null && $chart->HasChartAreaBorder ? $chart->ChartAreaBorderColor
            : ($chart !== null && $chart->RoundedCorners ? new SKColor(0x88, 0x88, 0x88) : self::BorderColor());
        $borderWidth = $chart !== null && $chart->HasChartAreaBorder
            ? (float)max(0.5, $chart->ChartAreaBorderWidthPt) : 1.0;
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $borderColor;
        $paint->StrokeWidth = $borderWidth;
        $paint->IsAntialias = false;
        if ($chart !== null && $chart->RoundedCorners) {
            $canvas->DrawRoundRect($area, 10.0, 10.0, $paint);
        } else {
            $canvas->DrawRect($area, $paint);
        }

        $paint->Dispose();
    }

    public function GetVisualBounds($area, $chart)
    {
        if ($chart === null) {
            return $area;
        }

        $this->chartCulture = $chart->Culture !== null ? $chart->Culture : $this->context->Culture;

        if ($chart->Kind === ChartKind::Pie) {
            return self::InsetAndClamp($area, $area->Left + $area->Width * 0.08, $area->Top + $area->Height * 0.06, $area->Right - $area->Width * 0.08, $area->Bottom - $area->Height * 0.08);
        }

        if ($chart->Kind === ChartKind::Radar) {
            return self::InsetAndClamp($area, $area->Left + $area->Width * 0.05, $area->Top + $area->Height * 0.08, $area->Right - $area->Width * 0.09, $area->Bottom - $area->Height * 0.12);
        }

        if (count($chart->Series) === 0) {
            return self::InsetAndClamp($area, $area->Left + $area->Width * 0.18, $area->Top + $area->Height * 0.18, $area->Right - $area->Width * 0.18, $area->Bottom - $area->Height * 0.18);
        }

        $plot = $this->ResolvePlotRect($area, $chart);
        $left = $plot->Left;
        $top = $plot->Top;
        $right = $plot->Right;
        $bottom = $plot->Bottom;

        $axisLabelMaxWidth = $this->MeasureMaxValueAxisLabelWidth($chart);
        $left = min($left, $plot->Left - 7.0 - $axisLabelMaxWidth);

        if ($chart->Title !== null && $chart->Title !== '') {
            $top = min($top, $area->Top + $area->Height * 0.06);
        }

        if (count($chart->Categories) > 0) {
            $bottom = max($bottom, $plot->Bottom + $this->LabelSizePt + 10.0);
        }

        $dateAxisLayout = $this->ResolveDateAxisLayout($chart, $plot->Width);
        if ($chart->HasDateCategoryAxis && abs($dateAxisLayout->RotationDeg) > 0.1) {
            $bottom = max($bottom, $plot->Bottom + 44.0);
        } elseif (!$chart->HasDateCategoryAxis && abs($this->ResolveStandardCategoryAxisRotation($chart, $plot->Width)) > 0.1) {
            $bottom = max($bottom, $plot->Bottom + 44.0);
        }

        if ($chart->LegendPosition !== null && count($chart->Series) > 0) {
            $legendBounds = $this->GetLegendBounds($area, $plot, $chart);
            $left = min($left, $legendBounds->Left);
            $right = max($right, $legendBounds->Right);
            $bottom = max($bottom, $legendBounds->Bottom);
        }

        return self::InsetAndClamp($area, $left - 2.0, $top - 2.0, $right + 2.0, $bottom + 2.0);
    }

    private function ResolvePlotRect($area, $chart)
    {
        if ($chart->Kind === ChartKind::Scatter) {
            return ScatterChartPlotLayout::Resolve($area, $chart);
        }

        if ($chart->Kind === ChartKind::Radar) {
            $radarLeft = $area->Left + $area->Width * 0.11;
            $radarTop = $area->Top + ($chart->Title !== null ? $area->Height * 0.1 : $area->Height * 0.1);
            $radarRight = $area->Right - $area->Width * 0.15;
            $radarBottom = $area->Bottom - $area->Height * ($chart->LegendPosition !== null ? 0.10 : 0.12);
            return new SKRect($radarLeft, $radarTop, $radarRight, $radarBottom);
        }

        if ($chart->HasManualPlot) {
            $l = $area->Left + (float)($chart->PlotX * $area->Width);
            $t = $area->Top + (float)($chart->PlotY * $area->Height);
            return new SKRect($l, $t, $l + (float)($chart->PlotW * $area->Width), $t + (float)($chart->PlotH * $area->Height));
        }

        $isPercentStackedArea = $chart->Kind === ChartKind::Area && $chart->IsPercentStacked;
        $axisLabelMaxWidth = $this->MeasureMaxValueAxisLabelWidth($chart);
        $leftPadding = max($area->Width * 0.06, $axisLabelMaxWidth + 14.0);
        $left = $area->Left + $leftPadding;
        $top = $area->Top + ($chart->Title !== null ? $area->Height * ($isPercentStackedArea ? 0.136 : 0.205) : $area->Height * 0.085);
        $right = $area->Right - $area->Width * 0.07;
        $bottomReserve = $isPercentStackedArea && $chart->LegendPosition !== null ? 0.17 : $this->BottomReserveFraction($chart, $area->Width);

        if ($chart->Kind === ChartKind::Column && !$chart->IsPercentStacked
            && !$chart->IsBar3D && $chart->Title !== null && $chart->LegendPosition !== null) {
            $left = $area->Left + max($area->Width * 0.06, $axisLabelMaxWidth + 14.5);
            $top = $area->Top + $area->Height * 0.185;
            $right = $area->Right - $area->Width * 0.031;
            $bottomReserve = 0.239;
        }

        if ($chart->Kind === ChartKind::Column && $chart->HasDataTable) {
            $left = $area->Left + $area->Width * 0.155;
            $top = $area->Top + $area->Height * 0.187;
            $right = $area->Right - $area->Width * 0.03;
            $bottomReserve = 0.575;
        }

        if ($chart->Kind === ChartKind::Column && $chart->IsPercentStacked
            && !$chart->IsBar3D && !$chart->HasDataTable) {
            $top = $area->Top + $area->Height * 0.136;
            $right = $area->Right - $area->Width * 0.02;
            $bottomReserve = 0.17;
        }

        if ($chart->Kind === ChartKind::Column && !$chart->IsBar3D
            && $chart->CategoryAxisTitle !== null && $chart->CategoryAxisTitle !== ''
            && $chart->ValueAxisTitle !== null && $chart->ValueAxisTitle !== '') {
            $left = $area->Left + max($area->Width * 0.06, $axisLabelMaxWidth + 14.5) + 22.0;
            $top = $area->Top + $area->Height * 0.172;
            $right = $area->Right - $area->Width * 0.0285;
            $bottomReserve = 0.324;
        }

        if ($chart->Kind === ChartKind::Area && $chart->IsArea3D && $chart->IsStacked
            && $chart->LegendPosition === 't') {
            $top = $area->Top + $area->Height * 0.34;
            $bottomReserve = 0.13;
        }

        $dateAxisLayout = $this->ResolveDateAxisLayout($chart, $area->Width * 0.83);
        if ($chart->HasDateCategoryAxis
            && abs($dateAxisLayout->RotationDeg) > 0.1
            && $chart->LegendPosition !== null) {
            // Excel keeps percent-stacked/date-axis line charts with rotated month labels a bit
            // tighter than our generic chart heuristic: the plot sits higher, is slightly
            // narrower, and leaves more room for the legend block below.
            $left = $area->Left + max($area->Width * 0.068, $axisLabelMaxWidth + 18.0);
            $top = $area->Top + ($chart->Title !== null ? $area->Height * 0.165 : $area->Height * 0.095);
            $right = $area->Right - $area->Width * 0.078;
            $bottomReserve = max($bottomReserve, 0.315);

            if ($chart->Kind === ChartKind::Line && $chart->IsStacked && !$chart->IsPercentStacked) {
                $left = $area->Left + max($area->Width * 0.064, $axisLabelMaxWidth + 17.0);
                $top = $area->Top + ($chart->Title !== null ? $area->Height * 0.172 : $area->Height * 0.1);
                $right = $area->Right - $area->Width * 0.064;
                $bottomReserve = max($bottomReserve, 0.323);
            }
        }

        $bottom = $area->Bottom - $area->Height * $bottomReserve;
        if ($chart->Kind === ChartKind::Column && !$chart->IsBar3D && $chart->TextSizePt !== null && !$chart->HasDataTable) {
            $textSize = (float)$chart->TextSizePt;
            $left = $area->Left + $axisLabelMaxWidth + $textSize * 1.5;
            $top = $area->Top + ($chart->Title === null ? $textSize : (float)$chart->TitleFont->Size + $textSize * 2.5);
            $right = $area->Right - $textSize * 1.5;
            $bottom = $area->Bottom - $textSize * ($chart->LegendPosition === 'b' ? 4.0 : 2.5);
        }
        if (self::UsesSideLegend($chart)) {
            $legendWidth = $this->SideLegendWidth($chart);
            if ($chart->LegendPosition === 'l') {
                $left += $legendWidth;
            } else {
                $right -= $legendWidth;
            }

            $bottom = $area->Bottom - (float)($chart->TextSizePt ?? 9.0) * 2.5;
        }
        if (!$chart->HasDateCategoryAxis && abs($this->ResolveStandardCategoryAxisRotation($chart, $right - $left)) < 0.1) {
            $bottom -= ChartCategoryLabels::ExtraHeight($chart, $this->context->GetFontContext($this->ChartFont($this->LabelSizePt, false)));
        }
        return LineChartAutomaticLayout::Plot($area, $chart, $axisLabelMaxWidth,
            $chart->HasDateCategoryAxis ? $dateAxisLayout->RotationDeg : $this->ResolveStandardCategoryAxisRotation($chart, $right - $left),
            new SKRect($left, $top, $right, $bottom));
    }

    private function DrawTitle($canvas, $page, $area, $chart, $centerY)
    {
        if ($chart->Title === null || $chart->Title === '') {
            return;
        }

        $clip = $this->ResolveTitleBounds($area, $chart, $centerY);
        $this->DrawTitleBorder($canvas, $clip, $chart);
        $titleFont = $this->ResolveTitleFont($chart);
        $titleText = $this->ResolveTitleText($chart);
        $this->DrawTitleText($canvas, $page, $clip, $titleText, $clip->MidX, $clip->MidY, $titleFont,
            $chart->TitleColor, TextAlign::Center, VAlign::Middle, $chart->TitleTextOutlineColor);
    }

    private function ResolveTitleBounds($area, $chart, $centerY)
    {
        $titleFont = $this->ResolveTitleFont($chart);
        $textWidth = $this->MeasureText($this->ResolveTitleText($chart), $titleFont);
        $horizontalPadding = max(3.5, $area->Width * 0.007);
        $verticalPadding = 0.25;
        if ($chart->Kind === ChartKind::Line && !$chart->IsLine3D && $chart->HasTitleBorder && !$chart->HasManualTitle) {
            $horizontalPadding += (float)$titleFont->Size * 0.05;
            $verticalPadding = 1.5;
        }
        $width = min($area->Width * ($chart->HasManualTitle ? 0.34 : 0.96), $textWidth + $horizontalPadding * 2.0);
        $left = $chart->HasManualTitle
            ? $area->Left + (float)($chart->TitleX * $area->Width)
            : $area->MidX - $width * 0.5;
        $right = $area->MidX + $width * 0.5;
        if ($chart->HasManualTitle) {
            $right = $left + $width;
        }

        $paint = $this->TextPaint($titleFont, $chart->TitleColor);
        $metrics = $paint->FontMetrics;
        if ($chart->HasManualTitle) {
            $manualTop = $area->Top + (float)($chart->TitleY * $area->Height);
            $bounds = new SKRect($left, $manualTop, $right, $manualTop + $metrics->Descent - $metrics->Ascent + $verticalPadding * 2.0);
            $paint->Dispose();
            return $bounds;
        }

        $baseline = $centerY - ($metrics->Ascent + $metrics->Descent) / 2.0;
        $top = $baseline + $metrics->Ascent - $verticalPadding;
        $bottom = $baseline + $metrics->Descent + $verticalPadding;
        $paint->Dispose();
        return new SKRect($left, $top, $right, $bottom);
    }

    private function DrawTitleBorder($canvas, $bounds, $chart)
    {
        if ($canvas === null || $chart === null || !$chart->HasTitleBorder) {
            return;
        }

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->TitleBorderColor;
        $paint->StrokeWidth = (float)max(0.5, $chart->TitleBorderWidthPt);
        $paint->IsAntialias = false;
        $canvas->DrawRect($bounds, $paint);
        $paint->Dispose();
    }

    private function DrawValueAxis($canvas, $page, $area, $plot, $chart, $min, $max, $major)
    {
        if ($max <= $min || $major <= 0.0) {
            return;
        }

        $major = self::ResolveRenderedMajorUnit($chart, $plot, $major);

        if ($chart->HasMinorValueGridlines) {
            $minor = self::ResolveRenderedMinorUnit($chart, $major);
            $gridlineRight = $this->ResolveValueGridlineRight($plot, $chart);
            if ($minor > 0.0 && $minor < $major) {
                $minorGrid = new SKPaint();
                $minorGrid->Style = SKPaintStyle::Stroke;
                $minorGrid->Color = $chart->MinorValueGridlineColor;
                $minorGrid->StrokeWidth = (float)max(0.5, $chart->MinorValueGridlineWidthPt);
                $minorGrid->IsAntialias = false;
                $epsilon = $minor * 0.1;
                for ($value = $min + $minor; $value < $max - $epsilon; $value += $minor) {
                    $ratio = $value / $major;
                    $nearestMajor = round($ratio);
                    if (abs($ratio - $nearestMajor) <= 0.001) {
                        continue;
                    }

                    $y = self::ValueToY($value, $min, $max, $plot);
                    $canvas->DrawLine($plot->Left, $y, $gridlineRight, $y, $minorGrid);
                }

                $minorGrid->Dispose();
            }
        }

        $grid = new SKPaint();
        $grid->Style = SKPaintStyle::Stroke;
        $grid->Color = $chart->ValueGridlineColor;
        $grid->StrokeWidth = (float)max(0.5, $chart->ValueGridlineWidthPt);
        $grid->IsAntialias = false;
        $gridlineRight = $this->ResolveValueGridlineRight($plot, $chart);
        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $y = self::ValueToY($value, $min, $max, $plot);

            if ($chart->HasValueGridlines) {
                $canvas->DrawLine($plot->Left, $y, $gridlineRight, $y, $grid);
            }

            $label = $this->FormatValue($value / $chart->ValueDisplayUnitDivisor, $chart->ValueFormatCode);
            $width = $this->MeasureText($label, $this->LabelSizePt);
            $clip = new SKRect($plot->Left - 8.0 - $width, $y - $this->LabelSizePt, $plot->Left - 1.0, $y + $this->LabelSizePt);
            if (!$chart->ScatterValueAxisDeleted) {
                $labelX = $chart->Kind === ChartKind::Scatter && $chart->LegendPosition === null
                    ? $plot->Left - 7.5 : $plot->Left - 5.0;
                $labelColor = $chart->AxisLabelColor;
                if ($value < 0.0 && $chart->ValueFormatCode !== null && $chart->ValueFormatCode !== ''
                    && stripos($chart->ValueFormatCode, '[Red]') !== false) {
                    $labelColor = SKColors::Red();
                }
                $this->DrawText($canvas, $page, $clip, $label, $labelX, $y, $this->LabelSizePt, false,
                    $labelColor, TextAlign::Right, VAlign::Middle, $chart->AxisTextOutlineColor);
            }
        }

        $grid->Dispose();

        ChartValueAxisLineRenderer::Draw($canvas, $plot, $chart, $min, $max, $major);
    }

    private function ResolveTitleText($chart)
    {
        if ($chart === null || $chart->Title === null || $chart->Title === '' || !$chart->TitleAllCaps) {
            return $chart !== null ? $chart->Title : '';
        }

        $culture = $chart->Culture ?? $this->chartCulture ?? CultureInfo::GetInvariantCulture();
        return strtoupper($chart->Title);
    }

    private function DrawCategoryAxis($canvas, $page, $area, $plot, $chart)
    {
        if ($chart->Kind === ChartKind::Line && !$chart->IsLine3D) {
            LineChartCategoryAxisRenderer::Draw($canvas, $plot, $chart, $this->dateSystem);
        } else {
            $axis = new SKPaint();
            $axis->Style = SKPaintStyle::Stroke;
            $axis->Color = self::GridColor();
            $axis->StrokeWidth = 1.0;
            $axis->IsAntialias = false;
            $canvas->DrawLine($plot->Left, $plot->Bottom, $plot->Right, $plot->Bottom, $axis);
            $axis->Dispose();
        }

        if ($chart->HasDateCategoryAxis) {
            $this->DrawDateCategoryAxis($canvas, $page, $area, $plot, $chart);
            return;
        }

        if ($this->suppressCategoryAxisLabels) {
            return;
        }

        $count = count($chart->Categories);
        $visibleSlice = $this->ResolveVisibleChartSlice($page, $area);
        $rotatedClip = new SKRect(
            max($plot->Left - 36.0, $visibleSlice->Left),
            $plot->Bottom + 2.0,
            min($plot->Right + 18.0, $visibleSlice->Right),
            $visibleSlice->Bottom);
        $rotationDeg = $this->ResolveStandardCategoryAxisRotation($chart, $plot->Width);
        for ($i = 0; $i < $count; $i++) {
            $x = self::CategoryX($i, $count, $plot, $chart->CrossBetween);
            $label = $chart->Categories[$i];
            if (abs($rotationDeg) > 0.1) {
                if ($x < $visibleSlice->Left - 10.0 || $x > $visibleSlice->Right + 24.0) {
                    continue;
                }

                $this->DrawRotatedText($canvas, $page, $rotatedClip, $label, $x + 2.0, $plot->Bottom + 6.0, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Right, VAlign::Top, $rotationDeg);
            } else {
                $lines = ChartCategoryLabels::Lines($label);
                $advance = ChartCategoryLabels::LineAdvance($this->context->GetFontContext($this->ChartFont($this->LabelSizePt, false)));
                for ($line = 0; $line < count($lines); $line++) {
                    $width = $this->MeasureText($lines[$line], $this->LabelSizePt);
                    $lineTop = $plot->Bottom + $line * $advance;
                    $clip = new SKRect($x - $width * 0.6, $lineTop + 2.0, $x + $width * 0.6,
                        min($visibleSlice->Bottom, $lineTop + $this->LabelSizePt * 2.2));
                    $this->DrawText($canvas, $page, $clip, $lines[$line], $x, $lineTop + 4.0, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Center, VAlign::Top);
                }
            }
        }
    }

    private function DrawScatterCategoryAxis($canvas, $page, $plot, $chart, $minimumY, $maximumY)
    {
        $minimum = 0.0;
        $maximum = 0.0;
        $major = 0.0;
        ScatterChartAxisMath::ResolveXRange($chart, $minimum, $maximum, $major);
        if ($major <= 0.0) {
            return;
        }

        $axisY = $chart->IsBubble && $minimumY < 0.0 && $maximumY > 0.0
            ? self::ValueToY(0.0, $minimumY, $maximumY, $plot) : $plot->Bottom;

        $axis = new SKPaint();
        $axis->Style = SKPaintStyle::Stroke;
        $axis->Color = $chart->CategoryAxisLineColor;
        $axis->StrokeWidth = (float)max(0.5, $chart->CategoryAxisLineWidthPt);
        $axis->IsAntialias = false;
        if ($chart->HasCategoryAxisLine) {
            $canvas->DrawLine($plot->Left, $axisY, $plot->Right, $axisY, $axis);
        }

        $count = (int)floor(($maximum - $minimum) / $major + 0.0001);
        if ($chart->HasCategoryGridlines) {
            $grid = new SKPaint();
            $grid->Style = SKPaintStyle::Stroke;
            $grid->Color = $chart->CategoryGridlineColor;
            $grid->StrokeWidth = (float)max(0.5, $chart->CategoryGridlineWidthPt);
            $grid->IsAntialias = false;
            for ($index = 0; $index <= $count; $index++) {
                $x = ScatterChartAxisMath::ToScreen($minimum + $index * $major, $minimum, $maximum, $plot->Left, $plot->Right);
                $canvas->DrawLine($x, $plot->Top, $x, $plot->Bottom, $grid);
            }

            $grid->Dispose();
        }
        for ($index = 0; $index <= $count; $index++) {
            $value = $minimum + $index * $major;
            $x = ScatterChartAxisMath::ToScreen($value, $minimum, $maximum, $plot->Left, $plot->Right);
            if ($chart->CategoryAxisMajorTickMark === 'out') {
                $canvas->DrawLine($x, $axisY, $x, $axisY + 3.0, $axis);
            }

            $labelValue = $value == 0.0 && ChartXmlParser::IsDateFormat($chart->ScatterXFormatCode)
                ? 2.0 : $value;
            $label = $this->FormatValue($labelValue, $chart->ScatterXFormatCode);
            $clip = new SKRect($x - 24.0, $axisY + 4.0, $x + 24.0, $axisY + 20.0);
            if (!$chart->ScatterCategoryAxisDeleted) {
                $this->DrawText($canvas, $page, $clip, $label, $x, $axisY + 6.0,
                    $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Center, VAlign::Top,
                    $chart->AxisTextOutlineColor);
            }
        }

        $axis->Dispose();
    }

    private function DrawScatterSeries($canvas, $plot, $chart, $minimumY, $maximumY)
    {
        $minimumX = 0.0;
        $maximumX = 0.0;
        $majorX = 0.0;
        ScatterChartAxisMath::ResolveXRange($chart, $minimumX, $maximumX, $majorX);
        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $series = $chart->Series[$seriesIndex];
            $points = array();
            $count = min(count($series->XValues), count($series->Values));
            for ($index = 0; $index < $count; $index++) {
                if ($series->XValues[$index] === null || $series->Values[$index] === null) {
                    continue;
                }

                $xValue = $series->XValues[$index];
                $yValue = $series->Values[$index];
                if ($xValue < $minimumX || $xValue > $maximumX || $yValue < $minimumY || $yValue > $maximumY) {
                    continue;
                }

                $points[] = new SKPoint(
                    ScatterChartAxisMath::ToScreen($xValue, $minimumX, $maximumX, $plot->Left, $plot->Right),
                    self::ValueToY($yValue, $minimumY, $maximumY, $plot));
            }

            if ($series->HasVisibleLine && count($points) > 1) {
                $path = new SKPath();
                $line = new SKPaint();
                $line->Style = SKPaintStyle::Stroke;
                $line->Color = $series->LineColor;
                $line->StrokeWidth = (float)$series->LineWidthPt;
                $line->IsAntialias = true;
                $path->MoveTo($points[0]);
                if ($series->SmoothLine) {
                    ScatterSmoothPathBuilder::Append($path, $points);
                } else {
                    for ($index = 1; $index < count($points); $index++) {
                        $path->LineTo($points[$index]);
                    }
                }
                $canvas->DrawPath($path, $line);
                $line->Dispose();
                $path->Dispose();
            }

            $this->DrawSeriesMarkers($canvas, $series, $points);
        }
    }

    private function DrawHiddenAxisScatter($canvas, $page, $area, $chart)
    {
        $plot = new SKRect($area->Left + $area->Width * 0.023, $area->Top + $area->Height * 0.04,
            $area->Right - $area->Width * 0.102, $area->Bottom - $area->Height * 0.04);
        $minimumY = 0.0;
        $maximumY = 0.0;
        $majorY = 0.0;
        self::ComputeValueAxis($chart, $minimumY, $maximumY, $majorY);
        $minimumX = 0.0;
        $maximumX = 0.0;
        $majorX = 0.0;
        ScatterChartAxisMath::ResolveXRange($chart, $minimumX, $maximumX, $majorX);
        $gridColor = SKColors::Black();
        $grid = new SKPaint();
        $grid->Style = SKPaintStyle::Stroke;
        $grid->Color = $gridColor;
        $grid->StrokeWidth = 0.7;
        $grid->IsAntialias = false;
        $xSteps = (int)round(($maximumX - $minimumX) / $majorX);
        for ($i = 0; $i <= $xSteps; $i++) {
            $x = ScatterChartAxisMath::ToScreen($minimumX + $i * $majorX, $minimumX, $maximumX, $plot->Left, $plot->Right);
            $canvas->DrawLine($x, $plot->Top, $x, $plot->Bottom, $grid);
        }
        $ySteps = (int)round(($maximumY - $minimumY) / $majorY);
        for ($i = 0; $i <= $ySteps; $i++) {
            $y = self::ValueToY($minimumY + $i * $majorY, $minimumY, $maximumY, $plot);
            $canvas->DrawLine($plot->Left, $y, $plot->Right, $y, $grid);
        }

        $grid->Dispose();

        $this->DrawText($canvas, $page, $area, $chart->Title, $area->MidX, $plot->Top + 3.0,
            10.0, true, SKColors::Black(), TextAlign::Center, VAlign::Middle);
        $this->DrawText($canvas, $page, $area, $chart->CategoryAxisTitle, $plot->MidX, $plot->Bottom - 6.0,
            10.0, true, SKColors::Black(), TextAlign::Center, VAlign::Middle);
        $this->DrawRotatedText($canvas, $page, $area, $chart->ValueAxisTitle, $plot->Left + 7.5, $plot->MidY,
            10.0, true, SKColors::Black(), TextAlign::Center, VAlign::Middle, -90.0);

        if (count($chart->Series) === 0) {
            return;
        }

        $series = $chart->Series[0];
        $colors = array(
            new SKColor(0x68, 0x84, 0x40), new SKColor(0x7e, 0x9a, 0x49),
            new SKColor(0x90, 0xac, 0x52), new SKColor(0xa2, 0xbd, 0x65),
            new SKColor(0xb5, 0xcd, 0x89), new SKColor(0xc7, 0xda, 0xad),
            new SKColor(0xd5, 0xe3, 0xc6));
        $count = min(count($series->XValues), count($series->Values));
        for ($i = 0; $i < $count; $i++) {
            if ($series->XValues[$i] === null || $series->Values[$i] === null) {
                continue;
            }
            $color = $colors[min($i, count($colors) - 1)];
            $x = ScatterChartAxisMath::ToScreen($series->XValues[$i], $minimumX, $maximumX, $plot->Left, $plot->Right);
            $y = self::ValueToY($series->Values[$i], $minimumY, $maximumY, $plot);
            $shadow = new SKPaint();
            $shadow->Style = SKPaintStyle::Fill;
            $shadow->Color = new SKColor(0, 0, 0, 45);
            $shadow->IsAntialias = true;
            $shadowBlur = SKMaskFilter::CreateBlur(SKBlurStyle::Normal, 1.4);
            $shadow->MaskFilter = $shadowBlur;
            $marker = new SKPaint();
            $marker->Style = SKPaintStyle::Fill;
            $marker->Color = $color;
            $marker->IsAntialias = true;
            $canvas->DrawCircle($x + 0.8, $y + 1.0, 4.7, $shadow);
            $canvas->DrawCircle($x, $y, 4.5, $marker);
            $legendY = $plot->Top + $plot->Height * (0.294 + $i * 0.0695);
            $legendX = $plot->Right + $area->Width * 0.039;
            $canvas->DrawCircle($legendX, $legendY, 3.5, $marker);
            $this->DrawText($canvas, $page, $area, $this->FormatValue($series->XValues[$i], $chart->ScatterXFormatCode),
                $legendX + 5.0, $legendY, 10.0, false, SKColors::Black(), TextAlign::Left, VAlign::Middle);
            $marker->Dispose();
            $shadow->Dispose();
            $shadowBlur->Dispose();
        }
    }

    private function DrawAxisTitles($canvas, $page, $area, $plot, $chart)
    {
        $axisTitleSizePt = 10.0;
        if ($chart->CategoryAxisTitle !== null && $chart->CategoryAxisTitle !== '') {
            $this->DrawText(
                $canvas,
                $page,
                $area,
                $chart->CategoryAxisTitle,
                $plot->MidX,
                $plot->Bottom + ($chart->Kind === ChartKind::Scatter ? 26.0 : 25.0),
                $axisTitleSizePt,
                false,
                $chart->AxisLabelColor,
                TextAlign::Center,
                VAlign::Middle);
        }

        if ($chart->ValueAxisTitle !== null && $chart->ValueAxisTitle !== '') {
            $this->DrawRotatedText(
                $canvas,
                $page,
                $area,
                $chart->ValueAxisTitle,
                $plot->Left - 54.0,
                $plot->MidY,
                $axisTitleSizePt,
                false,
                $chart->AxisLabelColor,
                TextAlign::Center,
                VAlign::Middle,
                -90.0);
        }
    }

    private function ResolveStandardCategoryAxisRotation($chart, $availableWidth)
    {
        if ($chart === null || $chart->HasDateCategoryAxis || abs($chart->CategoryAxisTextRotationDeg) <= 0.1) {
            return 0.0;
        }

        $count = $chart->Categories !== null ? count($chart->Categories) : 0;
        if ($count <= 0) {
            return 0.0;
        }

        $slotWidth = $availableWidth / $count;
        $widestLabel = 0.0;
        for ($i = 0; $i < count($chart->Categories); $i++) {
            $label = $chart->Categories[$i] ?? '';
            $widestLabel = max($widestLabel, $this->MeasureText($label, $this->LabelSizePt));
        }

        if ($widestLabel <= $slotWidth * 0.9) {
            return 0.0;
        }

        return (float)$chart->CategoryAxisTextRotationDeg;
    }

    private function DrawSeries($canvas, $page, $plot, $chart, $min, $max)
    {
        $pointCount = self::MaxPointCount($chart);

        if ($chart->Kind === ChartKind::Radar) {
            return;
        }

        if ($chart->Kind === ChartKind::Column) {
            $this->DrawColumns($canvas, $page, $plot, $chart, $min, $max, $pointCount);
            return;
        }

        if ($chart->Kind === ChartKind::Bar) {
            $this->DrawBars($canvas, $plot, $chart, $min, $max, $pointCount);
            return;
        }

        $baselineY = self::ValueToY(max($min, 0.0), $min, $max, $plot);
        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $series = $chart->Series[$seriesIndex];
            $displayValues = ChartStackingMath::BuildDisplayValues($chart, $seriesIndex);
            $renderPoints = array();
            $line = new SKPaint();
            $line->Style = SKPaintStyle::Stroke;
            $line->Color = $series->LineColor;
            $line->StrokeWidth = (float)$series->LineWidthPt;
            $line->IsAntialias = true;
            $line->StrokeCap = SKStrokeCap::Round;
            $line->StrokeJoin = SKStrokeJoin::Round;
            $path = new SKPath();
            $started = false;
            $datePoints = ChartSeriesPointBuilder::BuildDateSeriesPoints($chart, $displayValues, $plot, $min, $max, $this->dateSystem);
            if ($datePoints !== null) {
                for ($i = 0; $i < count($datePoints); $i++) {
                    $point = $datePoints[$i];
                    $renderPoints[] = $point;
                    if (!$started) {
                        $path->MoveTo($point->X, $point->Y);
                        $started = true;
                    } else {
                        $path->LineTo($point->X, $point->Y);
                    }
                }
            } else {
                for ($i = 0; $i < count($displayValues); $i++) {
                    if ($displayValues[$i] === null) {
                        $started = false;
                        continue;
                    }

                    $x = $this->PointX($chart, $i, $pointCount, $plot);
                    $y = self::ValueToY($displayValues[$i], $min, $max, $plot);
                    $renderPoints[] = new SKPoint($x, $y);
                    if (!$started) {
                        $path->MoveTo($x, $y);
                        $started = true;
                    } else {
                        $path->LineTo($x, $y);
                    }
                }
            }

            if ($chart->Kind === ChartKind::Area) {
                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = new SKColor($series->Color->Red, $series->Color->Green, $series->Color->Blue, 255);
                $fill->IsAntialias = true;
                $areaPath = $this->BuildAreaFillPath($chart, $plot, $min, $max, $pointCount, $seriesIndex, $displayValues, $renderPoints, $baselineY);
                if ($areaPath !== null) {
                    $canvas->DrawPath($areaPath, $fill);
                    $areaPath->Dispose();
                }

                $fill->Dispose();
            }

            if ($series->HasVisibleLine) {
                ChartSeriesStrokeRenderer::Draw($canvas, $plot, $chart, $path, $line);
            }
            $path->Dispose();
            $line->Dispose();

            $this->DrawSeriesMarkers($canvas, $series, $renderPoints);
        }
    }

    private function BuildAreaFillPath($chart, $plot, $min, $max, $pointCount, $seriesIndex, $upperValues, $upperPoints, $baselineY)
    {
        if ($upperPoints === null || count($upperPoints) === 0) {
            return null;
        }

        $areaPath = new SKPath();
        for ($i = 0; $i < count($upperPoints); $i++) {
            $point = $upperPoints[$i];
            if ($i === 0) {
                $areaPath->MoveTo($point->X, $point->Y);
            } else {
                $areaPath->LineTo($point->X, $point->Y);
            }
        }

        if ($chart !== null && ($chart->IsStacked || $chart->IsPercentStacked) && $seriesIndex > 0) {
            $lowerValues = ChartStackingMath::BuildDisplayValues($chart, $seriesIndex - 1);
            $lowerPoints = ChartSeriesPointBuilder::BuildDateSeriesPoints($chart, $lowerValues, $plot, $min, $max, $this->dateSystem);
            if ($lowerPoints === null) {
                $lowerPoints = $this->BuildCategorySeriesPoints($chart, $lowerValues, $pointCount, $plot, $min, $max);
            }
        } else {
            $lowerPoints = $this->BuildAreaBaselinePoints($chart, $upperValues, $upperPoints, $pointCount, $plot, $baselineY);
        }

        if ($lowerPoints === null || count($lowerPoints) === 0) {
            $areaPath->Dispose();
            return null;
        }

        for ($i = count($lowerPoints) - 1; $i >= 0; $i--) {
            $areaPath->LineTo($lowerPoints[$i]->X, $lowerPoints[$i]->Y);
        }

        $areaPath->Close();
        return $areaPath;
    }

    private function BuildCategorySeriesPoints($chart, $values, $pointCount, $plot, $min, $max)
    {
        $points = array();
        if ($values === null) {
            return $points;
        }

        for ($i = 0; $i < count($values); $i++) {
            if ($values[$i] === null) {
                continue;
            }

            $x = $this->PointX($chart, $i, $pointCount, $plot);
            $y = self::ValueToY($values[$i], $min, $max, $plot);
            $points[] = new SKPoint($x, $y);
        }

        return $points;
    }

    private function DrawRadar($canvas, $page, $area, $plot, $chart, $min, $max, $major)
    {
        if ($canvas === null || $chart === null || $chart->Series === null || count($chart->Series) === 0 || $max <= $min) {
            return;
        }

        $categoryCount = $chart->Categories !== null ? count($chart->Categories) : 0;
        if ($categoryCount <= 2) {
            return;
        }

        $major = self::ResolveRenderedMajorUnit($chart, $plot, $major);
        $center = new SKPoint($plot->MidX, $plot->MidY + 1.0);
        $radius = $this->ResolveRadarRadius($plot);

        $this->DrawRadarGrid($canvas, $chart, $center, $radius, $min, $max, $major);
        $this->DrawRadarCategoryLabels($canvas, $page, $area, $plot, $chart, $center, $radius);
        $this->DrawRadarValueAxisLabels($canvas, $page, $area, $plot, $chart, $center, $radius, $min, $max, $major);

        for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
            $series = $chart->Series[$seriesIndex];
            $points = $this->BuildRadarSeriesPoints($chart, $series, $center, $radius, $min, $max);
            if (count($points) === 0) {
                continue;
            }

            $path = $this->BuildRadarPath($points);
            if ($path === null) {
                continue;
            }

            if ($chart->RadarStyle === 'filled') {
                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = new SKColor($series->Color->Red, $series->Color->Green, $series->Color->Blue, 110);
                $fill->IsAntialias = true;
                $canvas->DrawPath($path, $fill);
                $fill->Dispose();
            }

            if ($series->HasVisibleLine) {
                $line = new SKPaint();
                $line->Style = SKPaintStyle::Stroke;
                $line->Color = $series->LineColor;
                $line->StrokeWidth = (float)$series->LineWidthPt;
                $line->IsAntialias = true;
                $line->StrokeCap = SKStrokeCap::Round;
                $line->StrokeJoin = SKStrokeJoin::Round;
                $canvas->DrawPath($path, $line);
                $line->Dispose();
            }

            $path->Dispose();

            $this->DrawSeriesMarkers($canvas, $series, $points);
        }
    }

    private function ResolveRadarRadius($plot)
    {
        $radius = min($plot->Width, $plot->Height) * 0.462;
        if ($radius < 12.0) {
            $radius = 12.0;
        }

        return $radius;
    }

    private function DrawRadarGrid($canvas, $chart, $center, $radius, $min, $max, $major)
    {
        $categoryCount = $chart->Categories !== null ? count($chart->Categories) : 0;
        if ($categoryCount <= 2) {
            return;
        }

        $spokePaint = new SKPaint();
        $spokePaint->Style = SKPaintStyle::Stroke;
        $spokePaint->Color = $chart->HasCategoryGridlines ? $chart->CategoryGridlineColor : self::GridColor();
        $spokePaint->StrokeWidth = 1.0;
        $spokePaint->IsAntialias = false;

        $ringPaint = new SKPaint();
        $ringPaint->Style = SKPaintStyle::Stroke;
        $ringPaint->Color = $chart->HasValueGridlines ? self::GridColor() : $chart->CategoryGridlineColor;
        $ringPaint->StrokeWidth = 1.0;
        $ringPaint->IsAntialias = false;

        for ($index = 0; $index < $categoryCount; $index++) {
            $angle = self::RadarAngle($index, $categoryCount);
            $outerPoint = self::RadarPoint($center, $radius, $angle);
            $canvas->DrawLine($center->X, $center->Y, $outerPoint->X, $outerPoint->Y, $spokePaint);
        }

        if ($major <= 0.0) {
            $major = ($max - $min) / 5.0;
        }

        if ($major <= 0.0) {
            $spokePaint->Dispose();
            $ringPaint->Dispose();
            return;
        }

        $steps = (int)round(($max - $min) / $major);
        if ($steps <= 0) {
            $steps = 1;
        }

        for ($step = 1; $step <= $steps; $step++) {
            $value = $min + $step * $major;
            if ($value > $max + $major * 0.001) {
                $value = $max;
            }

            $currentRadius = self::RadarRadiusForValue($value, $min, $max, $radius);
            $path = new SKPath();
            for ($index = 0; $index < $categoryCount; $index++) {
                $angle = self::RadarAngle($index, $categoryCount);
                $point = self::RadarPoint($center, $currentRadius, $angle);
                if ($index === 0) {
                    $path->MoveTo($point->X, $point->Y);
                } else {
                    $path->LineTo($point->X, $point->Y);
                }
            }

            $path->Close();
            $canvas->DrawPath($path, $ringPaint);
            $path->Dispose();
        }

        $spokePaint->Dispose();
        $ringPaint->Dispose();
    }

    private function DrawRadarCategoryLabels($canvas, $page, $area, $plot, $chart, $center, $radius)
    {
        if ($this->suppressCategoryAxisLabels || $chart->Categories === null) {
            return;
        }

        $count = count($chart->Categories);
        $labelRadius = $radius + 5.0;
        for ($index = 0; $index < $count; $index++) {
            $angle = self::RadarAngle($index, $count);
            $point = self::RadarPoint($center, $labelRadius, $angle);
            $label = $chart->Categories[$index] !== null ? $chart->Categories[$index] : '';
            $width = $this->MeasureText($label, $this->LabelSizePt);
            $clip = new SKRect($point->X - $width, $point->Y - $this->LabelSizePt * 1.1, $point->X + $width, $point->Y + $this->LabelSizePt * 1.1);
            $this->DrawText($canvas, $page, $clip, $label, $point->X, $point->Y, $this->LabelSizePt, false, $chart->AxisLabelColor, $this->RadarTextAlign($angle), $this->RadarVerticalAlign($angle));
        }
    }

    private function DrawRadarValueAxisLabels($canvas, $page, $area, $plot, $chart, $center, $radius, $min, $max, $major)
    {
        if ($max <= $min) {
            return;
        }

        if ($major <= 0.0) {
            $major = ($max - $min) / 5.0;
        }

        if ($major <= 0.0) {
            return;
        }

        $steps = (int)round(($max - $min) / $major);
        if ($steps <= 0) {
            $steps = 1;
        }

        for ($step = 0; $step <= $steps; $step++) {
            $value = $min + $step * $major;
            if ($step === $steps || $value > $max) {
                $value = $max;
            }

            $currentRadius = self::RadarRadiusForValue($value, $min, $max, $radius);
            $label = $this->FormatValue($value, $chart->ValueFormatCode);
            $width = $this->MeasureText($label, $this->LabelSizePt);
            $x = $center->X - 8.0;
            $y = $center->Y - $currentRadius;
            $clip = new SKRect($x - $width - 4.0, $y - $this->LabelSizePt, $x + 1.0, $y + $this->LabelSizePt);
            $this->DrawText($canvas, $page, $clip, $label, $x, $y, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Right, VAlign::Middle);
        }
    }

    private function BuildRadarSeriesPoints($chart, $series, $center, $radius, $min, $max)
    {
        $points = array();
        if ($chart === null || $series === null || $chart->Categories === null) {
            return $points;
        }

        $count = count($chart->Categories);
        for ($index = 0; $index < $count; $index++) {
            if ($index >= count($series->Values) || $series->Values[$index] === null) {
                continue;
            }

            $angle = self::RadarAngle($index, $count);
            $currentRadius = self::RadarRadiusForValue($series->Values[$index], $min, $max, $radius);
            $points[] = self::RadarPoint($center, $currentRadius, $angle);
        }

        return $points;
    }

    private function BuildRadarPath($points)
    {
        if ($points === null || count($points) === 0) {
            return null;
        }

        $path = new SKPath();
        for ($index = 0; $index < count($points); $index++) {
            if ($index === 0) {
                $path->MoveTo($points[$index]->X, $points[$index]->Y);
            } else {
                $path->LineTo($points[$index]->X, $points[$index]->Y);
            }
        }

        $path->Close();
        return $path;
    }

    private static function RadarAngle($index, $count)
    {
        if ($count <= 0) {
            return 0.0;
        }

        return -M_PI / 2.0 + M_PI * 2.0 * $index / $count;
    }

    private static function RadarPoint($center, $radius, $angle)
    {
        return new SKPoint($center->X + $radius * cos($angle), $center->Y + $radius * sin($angle));
    }

    private static function RadarRadiusForValue($value, $min, $max, $outerRadius)
    {
        if ($max <= $min) {
            return 0.0;
        }

        $ratio = ($value - $min) / ($max - $min);
        if ($ratio < 0.0) {
            $ratio = 0.0;
        } elseif ($ratio > 1.0) {
            $ratio = 1.0;
        }

        return $ratio * $outerRadius;
    }

    private function RadarTextAlign($angle)
    {
        $cosine = cos($angle);
        if ($cosine > 0.25) {
            return TextAlign::Left;
        }

        if ($cosine < -0.25) {
            return TextAlign::Right;
        }

        return TextAlign::Center;
    }

    private function RadarVerticalAlign($angle)
    {
        $sine = sin($angle);
        if ($sine > 0.35) {
            return VAlign::Top;
        }

        if ($sine < -0.35) {
            return VAlign::Baseline;
        }

        return VAlign::Middle;
    }

    private function BuildAreaBaselinePoints($chart, $values, $upperPoints, $pointCount, $plot, $baselineY)
    {
        $points = array();
        if ($upperPoints === null || count($upperPoints) === 0) {
            return $points;
        }

        if ($chart !== null && $chart->HasDateCategoryAxis && $chart->CategoryValues !== null) {
            $count = count($chart->CategoryValues);
            if ($values !== null && count($values) < $count) {
                $count = count($values);
            }

            for ($i = 0; $i < $count; $i++) {
                if ($chart->CategoryValues[$i] === null || $values === null || $values[$i] === null) {
                    continue;
                }

                $serial = ChartDateAxisMath::NormalizePointSerial($chart, $chart->CategoryValues[$i], $this->dateSystem);
                $plotMinDate = 0.0;
                $plotMaxDate = 0.0;
                if (!ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $this->dateSystem, $plotMinDate, $plotMaxDate)) {
                    break;
                }

                $points[] = new SKPoint(
                    ChartDateAxisMath::DateCategoryX($serial, $plotMinDate, $plotMaxDate, $plot),
                    $baselineY
                );
            }

            if (count($points) > 0) {
                return $points;
            }
        }

        for ($i = 0; $i < count($upperPoints); $i++) {
            $points[] = new SKPoint($upperPoints[$i]->X, $baselineY);
        }

        return $points;
    }

    private function DrawSeriesMarkers($canvas, $series, $points)
    {
        if ($canvas === null || $series === null || $points === null || count($points) === 0 || !$series->MarkerVisible) {
            return;
        }

        $diameter = max(3.0, $series->MarkerSize);
        $radius = $diameter * 0.5;
        $strokeWidth = (float)max(0.75, $series->MarkerStrokeWidthPt);

        $fill = new SKPaint();
        $fill->Style = SKPaintStyle::Fill;
        $fill->Color = $series->MarkerFillColor;
        $fill->IsAntialias = true;

        $stroke = new SKPaint();
        $stroke->Style = SKPaintStyle::Stroke;
        $stroke->Color = $series->MarkerStrokeColor;
        $stroke->StrokeWidth = $strokeWidth;
        $stroke->IsAntialias = true;

        for ($i = 0; $i < count($points); $i++) {
            $this->DrawMarker($canvas, $points[$i], $radius, $series->MarkerSymbol, $fill, $stroke);
        }

        $fill->Dispose();
        $stroke->Dispose();
    }

    private function DrawMarker($canvas, $center, $radius, $symbol, $fill, $stroke)
    {
        if ($symbol === 'none') {
            return;
        }

        if ($symbol === 'square') {
            $rect = new SKRect($center->X - $radius, $center->Y - $radius, $center->X + $radius, $center->Y + $radius);
            $canvas->DrawRect($rect, $fill);
            $canvas->DrawRect($rect, $stroke);
            return;
        }

        if ($symbol === 'diamond') {
            $path = new SKPath();
            $path->MoveTo($center->X, $center->Y - $radius);
            $path->LineTo($center->X + $radius, $center->Y);
            $path->LineTo($center->X, $center->Y + $radius);
            $path->LineTo($center->X - $radius, $center->Y);
            $path->Close();
            $canvas->DrawPath($path, $fill);
            $canvas->DrawPath($path, $stroke);
            $path->Dispose();
            return;
        }

        if ($symbol === 'triangle') {
            $path = new SKPath();
            $path->MoveTo($center->X, $center->Y - $radius);
            $path->LineTo($center->X + $radius * 0.9, $center->Y + $radius * 0.8);
            $path->LineTo($center->X - $radius * 0.9, $center->Y + $radius * 0.8);
            $path->Close();
            $canvas->DrawPath($path, $fill);
            $canvas->DrawPath($path, $stroke);
            $path->Dispose();
            return;
        }

        $canvas->DrawCircle($center, $radius, $fill);
        $canvas->DrawCircle($center, $radius, $stroke);
    }

    private function DrawColumns($canvas, $page, $plot, $chart, $min, $max, $pointCount)
    {
        if ($pointCount === 0 || count($chart->Series) === 0) {
            return;
        }

        $slot = $plot->Width / $pointCount;
        $columnGeometry = self::ResolveColumnGeometry($chart);
        $groupWidth = $slot * $columnGeometry[0];
        $barWidth = $slot * $columnGeometry[1];
        $seriesPitch = $slot * $columnGeometry[2];
        $baselineY = self::ValueToY(max($min, 0.0), $min, $max, $plot);

        if ($chart->IsStacked || $chart->IsPercentStacked) {
            $this->DrawStackedColumns($canvas, $page, $plot, $chart, $min, $max, $pointCount, $slot, $groupWidth, $baselineY);
            return;
        }

        for ($i = 0; $i < $pointCount; $i++) {
            $slotLeft = $plot->Left + $i * $slot + ($slot - $groupWidth) / 2.0;
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $y = self::ValueToY($series->Values[$i], $min, $max, $plot);
                $left = $slotLeft + $s * $seriesPitch;
                $top = min($y, $baselineY);
                $bottom = max($y, $baselineY);
                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = ChartExplicitFormatting::PointColor($series, $i);
                $fill->IsAntialias = true;
                $canvas->DrawRect(new SKRect($left, $top, $left + $barWidth, $bottom), $fill);
                $fill->Dispose();

                ChartBarBorder::Draw($canvas, new SKRect($left, $top, $left + $barWidth, $bottom), $series);
                $this->DrawColumnDataLabel($canvas, $page, $plot, $chart, $series, $series->Values[$i],
                    $left + $barWidth * 0.5, ($top + $bottom) * 0.5, $y);
                $this->DrawColumnErrorBars($canvas, $plot, $chart, $series, $s,
                    $left + $barWidth * 0.5, $series->Values[$i], $min, $max);
            }
        }
    }

    private function DrawColumnErrorBars($canvas, $plot, $chart, $series, $seriesIndex, $centerX, $endValue, $min, $max)
    {
        if ($series === null || !$series->HasErrorBars
            || $series->ErrorValueType !== 'stdErr') {
            return;
        }

        $standardError = self::ResolveSeriesStandardError($chart, $seriesIndex);
        if ($standardError <= 0.0) {
            return;
        }

        $top = self::ValueToY($endValue + $standardError, $min, $max, $plot);
        $bottom = self::ValueToY($endValue - $standardError, $min, $max, $plot);
        $capHalfWidth = 1.75;

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $series->ErrorBarColor;
        $paint->StrokeWidth = (float)max(0.5, $series->ErrorBarWidthPt);
        $paint->IsAntialias = false;

        $canvas->Save();
        $canvas->ClipRect($plot);
        $canvas->DrawLine($centerX, $top, $centerX, $bottom, $paint);
        if ($series->ErrorBarsHaveEndCaps) {
            $canvas->DrawLine($centerX - $capHalfWidth, $top, $centerX + $capHalfWidth, $top, $paint);
            $canvas->DrawLine($centerX - $capHalfWidth, $bottom, $centerX + $capHalfWidth, $bottom, $paint);
        }
        $canvas->Restore();

        $paint->Dispose();
    }

    private static function ResolveSeriesStandardError($chart, $seriesIndex)
    {
        if ($chart === null || $seriesIndex < 0 || $seriesIndex >= count($chart->Series)) {
            return 0.0;
        }

        $values = array();
        $series = $chart->Series[$seriesIndex];
        for ($pointIndex = 0; $pointIndex < count($series->Values); $pointIndex++) {
            if ($series->Values[$pointIndex] === null) {
                continue;
            }

            $value = $series->Values[$pointIndex];
            if ($chart->IsPercentStacked) {
                $total = 0.0;
                for ($otherIndex = 0; $otherIndex < count($chart->Series); $otherIndex++) {
                    $other = $chart->Series[$otherIndex];
                    if ($pointIndex < count($other->Values) && $other->Values[$pointIndex] !== null) {
                        $total += abs($other->Values[$pointIndex]);
                    }
                }

                if ($total <= 0.0) {
                    continue;
                }

                $value /= $total;
            }

            $values[] = $value;
        }

        if (count($values) < 2) {
            return 0.0;
        }

        $mean = 0.0;
        for ($index = 0; $index < count($values); $index++) {
            $mean += $values[$index];
        }
        $mean /= count($values);

        $squaredDifferences = 0.0;
        for ($index = 0; $index < count($values); $index++) {
            $difference = $values[$index] - $mean;
            $squaredDifferences += $difference * $difference;
        }

        $sampleDeviation = sqrt($squaredDifferences / (count($values) - 1));
        return $sampleDeviation / sqrt(count($values));
    }

    private function DrawDataTable($canvas, $page, $area, $plot, $chart)
    {
        if ($chart === null || !$chart->HasDataTable || count($chart->Series) === 0) {
            return;
        }

        $categoryCount = $chart->Categories !== null ? count($chart->Categories) : 0;
        if ($categoryCount === 0) {
            $categoryCount = self::MaxPointCount($chart);
        }
        if ($categoryCount === 0) {
            return;
        }

        $labelWidth = min($area->Width * 0.17, 49.0);
        $tableLeft = $plot->Left - $labelWidth;
        $tableTop = $plot->Bottom + 24.5;
        $rowHeight = 13.4;
        $headerBottom = $tableTop + $rowHeight;
        $tableBottom = $headerBottom + count($chart->Series) * $rowHeight;
        $cellWidth = $plot->Width / $categoryCount;

        $border = new SKPaint();
        $border->Style = SKPaintStyle::Stroke;
        $border->Color = $chart->DataTableLineColor;
        $border->StrokeWidth = (float)max(0.5, $chart->DataTableLineWidthPt);
        $border->IsAntialias = false;

        if ($chart->DataTableShowHorizontalBorders) {
            $canvas->DrawLine($plot->Left, $tableTop, $plot->Right, $tableTop, $border);
            for ($row = 0; $row <= count($chart->Series); $row++) {
                $y = $headerBottom + $row * $rowHeight;
                $canvas->DrawLine($tableLeft, $y, $plot->Right, $y, $border);
            }
        }

        if ($chart->DataTableShowVerticalBorders) {
            $canvas->DrawLine($tableLeft, $headerBottom, $tableLeft, $tableBottom, $border);
            for ($column = 0; $column <= $categoryCount; $column++) {
                $x = $plot->Left + $column * $cellWidth;
                $canvas->DrawLine($x, $tableTop, $x, $tableBottom, $border);
            }
        }

        if ($chart->DataTableShowOutline) {
            $canvas->DrawLine($plot->Left, $tableTop, $plot->Right, $tableTop, $border);
            $canvas->DrawLine($tableLeft, $tableBottom, $plot->Right, $tableBottom, $border);
            $canvas->DrawLine($plot->Right, $tableTop, $plot->Right, $tableBottom, $border);
        }

        $border->Dispose();

        for ($column = 0; $column < $categoryCount; $column++) {
            $category = $this->ResolveCategoryLabel($chart, $column);
            $centerX = $plot->Left + ($column + 0.5) * $cellWidth;
            $this->DrawText($canvas, $page, new SKRect($plot->Left + $column * $cellWidth, $tableTop,
                $plot->Left + ($column + 1) * $cellWidth, $headerBottom), $category, $centerX,
                $tableTop + $rowHeight * 0.5, $this->LabelSizePt, false, $chart->DataTableTextColor,
                TextAlign::Center, VAlign::Middle);
        }

        for ($row = 0; $row < count($chart->Series); $row++) {
            $seriesIndex = count($chart->Series) - 1 - $row;
            $series = $chart->Series[$seriesIndex];
            $centerY = $headerBottom + ($row + 0.5) * $rowHeight;
            $textLeft = $tableLeft + 3.0;
            if ($chart->DataTableShowKeys) {
                $key = new SKPaint();
                $key->Style = SKPaintStyle::Fill;
                $key->Color = $series->Color;
                $key->IsAntialias = false;
                $canvas->DrawRect(new SKRect($tableLeft + 3.0, $centerY - 2.3,
                    $tableLeft + 7.6, $centerY + 2.3), $key);
                $key->Dispose();
                $textLeft = $tableLeft + 10.0;
            }

            $this->DrawText($canvas, $page, new SKRect($tableLeft, $centerY - $rowHeight * 0.5,
                $plot->Left, $centerY + $rowHeight * 0.5), $series->Name, $textLeft, $centerY,
                $this->LabelSizePt, false, $chart->DataTableTextColor, TextAlign::Left, VAlign::Middle);

            for ($column = 0; $column < $categoryCount; $column++) {
                if ($column >= count($series->Values) || $series->Values[$column] === null) {
                    continue;
                }

                $formatCode = $series->ValueFormatCode !== null && $series->ValueFormatCode !== ''
                    ? $series->ValueFormatCode
                    : $chart->ValueFormatCode;
                $value = $this->FormatValue($series->Values[$column], $formatCode);
                $centerX = $plot->Left + ($column + 0.5) * $cellWidth;
                $this->DrawText($canvas, $page, new SKRect($plot->Left + $column * $cellWidth,
                    $centerY - $rowHeight * 0.5, $plot->Left + ($column + 1) * $cellWidth,
                    $centerY + $rowHeight * 0.5), $value, $centerX, $centerY, $this->LabelSizePt,
                    false, $chart->DataTableTextColor, TextAlign::Center, VAlign::Middle);
            }
        }
    }

    private function DrawBars($canvas, $plot, $chart, $min, $max, $pointCount)
    {
        if ($pointCount === 0 || count($chart->Series) === 0) {
            return;
        }

        $slot = $plot->Height / $pointCount;
        $barGeometry = self::ResolveColumnGeometry($chart);
        $groupHeight = $slot * $barGeometry[0];
        $barHeight = $slot * $barGeometry[1];
        $seriesPitch = $slot * $barGeometry[2];
        $baselineX = self::ValueToX(max($min, 0.0), $min, $max, $plot);

        if ($chart->IsStacked || $chart->IsPercentStacked) {
            $this->DrawStackedBars($canvas, $plot, $chart, $min, $max, $pointCount, $slot, $groupHeight, $baselineX);
            return;
        }

        for ($i = 0; $i < $pointCount; $i++) {
            $categoryIndex = $this->ResolveBarCategoryIndex($chart, $i, $pointCount);
            $slotTop = $plot->Top + $categoryIndex * $slot + ($slot - $groupHeight) / 2.0;
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $x = self::ValueToX($series->Values[$i], $min, $max, $plot);
                $top = $slotTop + (count($chart->Series) - 1 - $s) * $seriesPitch;
                $left = min($x, $baselineX);
                $right = max($x, $baselineX);
                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = $series->Color;
                $fill->IsAntialias = true;
                $canvas->DrawRect(new SKRect($left, $top, $right, $top + $barHeight), $fill);
                $fill->Dispose();

                ChartBarBorder::Draw($canvas, new SKRect($left, $top, $right, $top + $barHeight), $series);
            }
        }
    }

    private function DrawStackedBars($canvas, $plot, $chart, $min, $max, $pointCount, $slot, $barHeight, $baselineX)
    {
        $positiveTotals = array_fill(0, $pointCount, 0.0);
        $negativeTotals = array_fill(0, $pointCount, 0.0);

        if ($chart->IsPercentStacked) {
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                $count = min($pointCount, count($series->Values));
                for ($i = 0; $i < $count; $i++) {
                    if ($series->Values[$i] === null) {
                        continue;
                    }

                    $value = $series->Values[$i];
                    if ($value >= 0.0) {
                        $positiveTotals[$i] += $value;
                    } else {
                        $negativeTotals[$i] += abs($value);
                    }
                }
            }
        }

        $positiveStarts = array_fill(0, $pointCount, 0.0);
        $negativeStarts = array_fill(0, $pointCount, 0.0);
        for ($i = 0; $i < $pointCount; $i++) {
            $categoryIndex = $this->ResolveBarCategoryIndex($chart, $i, $pointCount);
            $top = $plot->Top + $categoryIndex * $slot + ($slot - $barHeight) / 2.0;
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $value = $series->Values[$i];
                if ($chart->IsPercentStacked) {
                    if ($value >= 0.0) {
                        if ($positiveTotals[$i] <= 0.0) {
                            continue;
                        }

                        $value /= $positiveTotals[$i];
                    } else {
                        if ($negativeTotals[$i] <= 0.0) {
                            continue;
                        }

                        $value /= $negativeTotals[$i];
                    }
                }

                if ($value >= 0.0) {
                    $startValue = $positiveStarts[$i];
                    $endValue = $startValue + $value;
                    $positiveStarts[$i] = $endValue;
                } else {
                    $startValue = $negativeStarts[$i];
                    $endValue = $startValue + $value;
                    $negativeStarts[$i] = $endValue;
                }

                $startX = self::ValueToX($startValue, $min, $max, $plot);
                $endX = self::ValueToX($endValue, $min, $max, $plot);
                $left = min($startX, $endX);
                $right = max($startX, $endX);

                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = $series->Color;
                $fill->IsAntialias = true;
                $canvas->DrawRect(new SKRect($left, $top, $right, $top + $barHeight), $fill);
                $fill->Dispose();

                ChartBarBorder::Draw($canvas, new SKRect($left, $top, $right, $top + $barHeight), $series);
            }
        }
    }

    private function DrawStackedColumns($canvas, $page, $plot, $chart, $min, $max, $pointCount, $slot, $barWidth, $baselineY)
    {
        $positiveTotals = array_fill(0, $pointCount, 0.0);
        $negativeTotals = array_fill(0, $pointCount, 0.0);

        if ($chart->IsPercentStacked) {
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                $count = min($pointCount, count($series->Values));
                for ($i = 0; $i < $count; $i++) {
                    if ($series->Values[$i] === null) {
                        continue;
                    }

                    $value = $series->Values[$i];
                    if ($value >= 0.0) {
                        $positiveTotals[$i] += $value;
                    } else {
                        $negativeTotals[$i] += abs($value);
                    }
                }
            }
        }

        $positiveStarts = array_fill(0, $pointCount, 0.0);
        $negativeStarts = array_fill(0, $pointCount, 0.0);
        for ($i = 0; $i < $pointCount; $i++) {
            $left = $plot->Left + $i * $slot + ($slot - $barWidth) / 2.0;
            for ($s = 0; $s < count($chart->Series); $s++) {
                $series = $chart->Series[$s];
                if ($i >= count($series->Values) || $series->Values[$i] === null) {
                    continue;
                }

                $value = $series->Values[$i];
                if ($chart->IsPercentStacked) {
                    if ($value >= 0.0) {
                        if ($positiveTotals[$i] <= 0.0) {
                            continue;
                        }

                        $value /= $positiveTotals[$i];
                    } else {
                        if ($negativeTotals[$i] <= 0.0) {
                            continue;
                        }

                        $value /= $negativeTotals[$i];
                    }
                }

                if ($value >= 0.0) {
                    $startValue = $positiveStarts[$i];
                    $endValue = $startValue + $value;
                    $positiveStarts[$i] = $endValue;
                } else {
                    $startValue = $negativeStarts[$i];
                    $endValue = $startValue + $value;
                    $negativeStarts[$i] = $endValue;
                }

                $startY = self::ValueToY($startValue, $min, $max, $plot);
                $endY = self::ValueToY($endValue, $min, $max, $plot);
                $top = min($startY, $endY);
                $bottom = max($startY, $endY);

                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = ChartExplicitFormatting::PointColor($series, $i);
                $fill->IsAntialias = true;
                $canvas->DrawRect(new SKRect($left, $top, $left + $barWidth, $bottom), $fill);
                $fill->Dispose();

                ChartBarBorder::Draw($canvas, new SKRect($left, $top, $left + $barWidth, $bottom), $series);
                $this->DrawColumnDataLabel($canvas, $page, $plot, $chart, $series, $series->Values[$i],
                    $left + $barWidth * 0.5, ($top + $bottom) * 0.5, $endY);
                $this->DrawColumnErrorBars($canvas, $plot, $chart, $series, $s,
                    $left + $barWidth * 0.5, $endValue, $min, $max);
            }
        }
    }

    private function DrawColumnDataLabel($canvas, $page, $plot, $chart, $series, $value, $centerX, $centerY, $endY)
    {
        if ($series === null || !$series->ShowValueDataLabels) {
            return;
        }

        if (($series->DataLabelPosition !== null && $series->DataLabelPosition !== '')
            && $series->DataLabelPosition !== 'ctr'
            && $series->DataLabelPosition !== 'inEnd') {
            return;
        }

        $formatCode = $series->ValueFormatCode !== null && $series->ValueFormatCode !== ''
            ? $series->ValueFormatCode
            : $chart->ValueFormatCode;
        $label = $this->FormatValue($value, $formatCode);
        $size = $series->DataLabelSizePt !== null ? (float)$series->DataLabelSizePt : $this->LabelSizePt;
        if ($series->DataLabelPosition === 'inEnd') {
            $centerY = $endY + ($value < 0.0 ? -1.0 : 1.0) * ($size * 0.65 + 2.0);
        }
        $this->DrawText($canvas, $page, $plot, $label, $centerX, $centerY, $size, false,
            $series->DataLabelColor, TextAlign::Center, VAlign::Middle);
    }

    private static function ResolveColumnGeometry($chart)
    {
        $seriesCount = $chart !== null && $chart->Series !== null ? count($chart->Series) : 0;
        if ($seriesCount <= 0) {
            return array(0.0, 0.0, 0.0);
        }

        if ($chart === null) {
            $legacyGroupWidth = 0.8;
            $legacyBarWidth = $legacyGroupWidth / $seriesCount;
            return array($legacyGroupWidth, $legacyBarWidth, $legacyBarWidth);
        }

        $gapRatio = (float)($chart->GapWidthPercent / 100.0);
        if ($gapRatio < 0.0) {
            $gapRatio = 0.0;
        }

        if ($chart->IsStacked || $chart->IsPercentStacked) {
            $stackedWidthRatio = 1.0 / (1.0 + $gapRatio);
            return array($stackedWidthRatio, $stackedWidthRatio, 0.0);
        }

        $overlapRatio = (float)($chart->OverlapPercent / 100.0);
        if ($overlapRatio < -1.0) {
            $overlapRatio = -1.0;
        } elseif ($overlapRatio > 1.0) {
            $overlapRatio = 1.0;
        }

        $pitchFactor = 1.0 - $overlapRatio;
        if ($pitchFactor < 0.0) {
            $pitchFactor = 0.0;
        }

        $spanFactor = 1.0 + ($seriesCount - 1) * $pitchFactor;
        $denominator = $spanFactor + $gapRatio;
        if ($denominator <= 0.0) {
            $fallbackGroupWidth = 0.8;
            $fallbackBarWidth = $fallbackGroupWidth / $seriesCount;
            return array($fallbackGroupWidth, $fallbackBarWidth, $fallbackBarWidth);
        }

        $barWidthRatio = 1.0 / $denominator;
        $seriesPitchRatio = $barWidthRatio * $pitchFactor;
        $groupWidthRatio = $barWidthRatio * $spanFactor;
        return array($groupWidthRatio, $barWidthRatio, $seriesPitchRatio);
    }

    private function DrawLegend($canvas, $page, $area, $plot, $chart)
    {
        $legendItemCount = $this->LegendItemCount($chart);
        if ($chart->LegendPosition === null || $legendItemCount === 0) {
            return;
        }

        $sideLegend = self::UsesSideLegend($chart);
        $y = $this->LegendCenterY($area, $plot, $chart);
        $swatch = $this->ResolveLegendSwatchWidth($chart);
        $gap = $this->ResolveLegendTextGap($chart);
        $itemGap = $this->ResolveLegendItemGap($chart);

        $widths = array_fill(0, $legendItemCount, 0.0);
        $total = 0.0;
        for ($i = 0; $i < $legendItemCount; $i++) {
            $widths[$i] = $swatch + $gap + $this->MeasureText($this->LegendItemLabel($chart, $i), $this->LabelSizePt);
            $total += $widths[$i] + ($i > 0 ? $itemGap : 0.0);
        }

        $x = $this->ResolveLegendLeft($area, $chart, $total);
        $legendTextBottom = $y;
        $legendSwatchHeight = $chart->HasManualLegend ? 4.5 : ($this->IsCompactColumnLegend($chart) ? 5.0 : 5.5);
        $legendSwatchVerticalOffset = $chart->HasManualLegend ? -1.0 : ($this->IsCompactColumnLegend($chart) ? -0.5 : 0.0);

        $metricsPaint = $this->TextPaint($this->ChartFont($this->LabelSizePt, false), $chart->LegendTextColor);
        $metrics = $metricsPaint->FontMetrics;
        $legendTextBottom = $y + ($metrics->Descent - $metrics->Ascent) / 2.0 - 1.4;
        $metricsPaint->Dispose();

        $lineHeight = $this->LabelSizePt * 1.7;
        if ($sideLegend) {
            $textBottomOffset = $legendTextBottom - $y;
            $x = $chart->LegendPosition === 'l' ? $area->Left + $this->LabelSizePt : $area->Right - $this->SideLegendWidth($chart);
            $y = $area->MidY - $lineHeight * ($legendItemCount - 1) * 0.5;
            $legendTextBottom = $y + $textBottomOffset;
        }

        for ($i = 0; $i < $legendItemCount; $i++) {
            $label = $this->LegendItemLabel($chart, $i);
            $color = $this->LegendItemColor($chart, $i);
            if ($chart->Kind === ChartKind::Column || $chart->Kind === ChartKind::Bar || $chart->Kind === ChartKind::Pie
                || $chart->Kind === ChartKind::Area || ($chart->Kind === ChartKind::Line && $chart->IsLine3D)) {
                $rectBottom = $legendTextBottom + $legendSwatchVerticalOffset;
                $rectTop = $rectBottom - $legendSwatchHeight;
                $swatchRect = new SKRect($x, $rectTop, $x + $legendSwatchHeight, $rectBottom);
                $patternDrawn = false;
                if ($chart->Kind === ChartKind::Area) {
                    $seriesIndex = $this->ResolveLegendSeriesIndex($chart, $i);
                    if ($seriesIndex >= 0 && $seriesIndex < count($chart->Series)) {
                        $patternDrawn = ChartPatternRenderer::DrawRect($canvas, $swatchRect, $chart->Series[$seriesIndex], $color);
                    }
                }

                if (!$patternDrawn) {
                    $fill = new SKPaint();
                    $fill->Style = SKPaintStyle::Fill;
                    $fill->Color = $color;
                    $fill->IsAntialias = true;
                    $canvas->DrawRect($swatchRect, $fill);
                    $fill->Dispose();
                }
            } elseif ($chart->IsBubble) {
                $legendSeriesIndex = $this->ResolveLegendSeriesIndex($chart, $i);
                if ($legendSeriesIndex >= 0 && $legendSeriesIndex < count($chart->Series)
                    && $chart->Series[$legendSeriesIndex]->Bubble3D) {
                    BubbleSphereRenderer::Draw($canvas, new SKPoint($x + $swatch * 0.5, $y), 2.45, $color);
                } else {
                    $fill = new SKPaint();
                    $fill->Style = SKPaintStyle::Fill;
                    $fill->Color = $color;
                    $fill->IsAntialias = true;
                    $canvas->DrawCircle($x + $swatch * 0.5, $y, 2.45, $fill);
                    $fill->Dispose();
                }
            } elseif ($chart->Kind === ChartKind::Radar && $chart->RadarStyle === 'filled') {
                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = new SKColor($color->Red, $color->Green, $color->Blue, 110);
                $fill->IsAntialias = true;

                $stroke = new SKPaint();
                $stroke->Style = SKPaintStyle::Stroke;
                $stroke->Color = $color;
                $stroke->StrokeWidth = 1.4;
                $stroke->IsAntialias = true;

                $rectBottom = $legendTextBottom;
                $rectTop = $rectBottom - $legendSwatchHeight;
                $rect = new SKRect($x, $rectTop, $x + $legendSwatchHeight, $rectBottom);
                $canvas->DrawRect($rect, $fill);
                $canvas->DrawRect($rect, $stroke);

                $fill->Dispose();
                $stroke->Dispose();
            } else {
                $legendSeriesIndex = $this->ResolveLegendSeriesIndex($chart, $i);
                $showLine = $chart->Kind !== ChartKind::Scatter
                    || ($legendSeriesIndex >= 0 && $legendSeriesIndex < count($chart->Series) && $chart->Series[$legendSeriesIndex]->HasVisibleLine);
                if ($showLine) {
                    $legendLineWidth = LineChartDecorationLayout::UsesAutomaticBottomLegend($chart) ? 2.25 : 2.5;
                    if ($chart->Kind === ChartKind::Scatter && $legendSeriesIndex >= 0 && $legendSeriesIndex < count($chart->Series)) {
                        $legendLineWidth = (float)$chart->Series[$legendSeriesIndex]->LineWidthPt;
                    }
                    $line = new SKPaint();
                    $line->Style = SKPaintStyle::Stroke;
                    $line->Color = $color;
                    $line->StrokeWidth = $legendLineWidth;
                    $line->IsAntialias = true;
                    $line->StrokeCap = SKStrokeCap::Round;
                    $canvas->DrawLine($x, $y, $x + $swatch, $y, $line);
                    $line->Dispose();
                }

                if ($chart->Kind === ChartKind::Line || $chart->Kind === ChartKind::Scatter) {
                    if ($legendSeriesIndex >= 0 && $legendSeriesIndex < count($chart->Series) && $chart->Series[$legendSeriesIndex]->MarkerVisible) {
                        $series = $chart->Series[$legendSeriesIndex];

                        $fill = new SKPaint();
                        $fill->Style = SKPaintStyle::Fill;
                        $fill->Color = $series->MarkerFillColor;
                        $fill->IsAntialias = true;

                        $stroke = new SKPaint();
                        $stroke->Style = SKPaintStyle::Stroke;
                        $stroke->Color = $series->MarkerStrokeColor;
                        $stroke->StrokeWidth = (float)max(0.75, $series->MarkerStrokeWidthPt);
                        $stroke->IsAntialias = true;

                        $this->DrawMarker($canvas, new SKPoint($x + $swatch * 0.5, $y),
                            max(3.0, $series->MarkerSize) * 0.5, $series->MarkerSymbol, $fill, $stroke);

                        $fill->Dispose();
                        $stroke->Dispose();
                    }
                }

                if ($chart->Kind === ChartKind::Radar && $chart->RadarStyle === 'marker') {
                    $seriesIndex = $this->ResolveLegendSeriesIndex($chart, $i);
                    if ($chart->Series !== null && $seriesIndex >= 0 && $seriesIndex < count($chart->Series)) {
                        $fill = new SKPaint();
                        $fill->Style = SKPaintStyle::Fill;
                        $fill->Color = $chart->Series[$seriesIndex]->MarkerFillColor;
                        $fill->IsAntialias = true;

                        $stroke = new SKPaint();
                        $stroke->Style = SKPaintStyle::Stroke;
                        $stroke->Color = $chart->Series[$seriesIndex]->MarkerStrokeColor;
                        $stroke->StrokeWidth = 1.0;
                        $stroke->IsAntialias = true;

                        $this->DrawMarker($canvas, new SKPoint($x + $swatch * 0.5, $y), 2.35, $chart->Series[$seriesIndex]->MarkerSymbol, $fill, $stroke);

                        $fill->Dispose();
                        $stroke->Dispose();
                    }
                }
            }
            $labelWidth = $this->MeasureText($label, $this->LabelSizePt);
            $clip = new SKRect($x + $swatch + $gap, $y - $this->LabelSizePt, $x + $swatch + $gap + $labelWidth + 2.0, $y + $this->LabelSizePt);
            $this->DrawText($canvas, $page, $clip, $label, $x + $swatch + $gap, $y, $this->LabelSizePt, false,
                $chart->LegendTextColor, TextAlign::Left, VAlign::Middle, $chart->LegendTextOutlineColor);
            if ($sideLegend) {
                $y += $lineHeight;
                $legendTextBottom += $lineHeight;
            } else {
                $x += $widths[$i] + $itemGap;
            }
        }
    }

    private function GetLegendBounds($area, $plot, $chart)
    {
        if (self::UsesSideLegend($chart)) {
            $width = $this->SideLegendWidth($chart);
            $left = $chart->LegendPosition === 'l' ? $area->Left + $this->LabelSizePt : $area->Right - $width;
            $height = $this->LegendItemCount($chart) * $this->LabelSizePt * 1.7;
            return new SKRect($left, $area->MidY - $height * 0.5, $left + $width, $area->MidY + $height * 0.5);
        }

        $y = $this->LegendCenterY($area, $plot, $chart);
        $swatch = $this->ResolveLegendSwatchWidth($chart);
        $gap = $this->ResolveLegendTextGap($chart);
        $itemGap = $this->ResolveLegendItemGap($chart);

        $total = 0.0;
        $legendItemCount = $this->LegendItemCount($chart);
        for ($i = 0; $i < $legendItemCount; $i++) {
            $total += $swatch + $gap + $this->MeasureText($this->LegendItemLabel($chart, $i), $this->LabelSizePt);
            if ($i > 0) {
                $total += $itemGap;
            }
        }

        $x = $this->ResolveLegendLeft($area, $chart, $total);
        return new SKRect($x, $y - 8.0, $x + $total, $y + 8.0);
    }

    private function BottomReserveFraction($chart, $availableWidth)
    {
        $bottomReserve = $chart->LegendPosition !== null ? 0.2 : 0.12;
        $dateAxisLayout = $this->ResolveDateAxisLayout($chart, $availableWidth * 0.83);
        if ($chart->HasDateCategoryAxis && abs($dateAxisLayout->RotationDeg) > 0.1) {
            if ($chart->LegendPosition !== null) {
                $bottomReserve = 0.285;
            } else {
                $bottomReserve = 0.205;
            }
        }

        return $bottomReserve;
    }

    private function ResolveLegendSeriesIndex($chart, $legendIndex)
    {
        if ($chart === null || $chart->Series === null || $legendIndex < 0 || $legendIndex >= count($chart->Series)) {
            return $legendIndex;
        }

        if ($chart->Kind === ChartKind::Bar && !$chart->IsStacked && !$chart->IsPercentStacked) {
            return count($chart->Series) - 1 - $legendIndex;
        }

        if ($chart->Kind === ChartKind::Column && ($chart->IsStacked || $chart->IsPercentStacked) && self::UsesSideLegend($chart)) {
            return count($chart->Series) - 1 - $legendIndex;
        }

        return $legendIndex;
    }

    private static function UsesSideLegend($chart)
    {
        return !$chart->HasManualLegend && !$chart->IsBar3D
            && ($chart->Kind === ChartKind::Column || $chart->Kind === ChartKind::Bar)
            && ($chart->LegendPosition === 'r' || $chart->LegendPosition === 'l' || $chart->LegendPosition === 'tr');
    }

    private function SideLegendWidth($chart)
    {
        $textWidth = 0.0;
        $size = $chart->TextSizePt !== null ? (float)$chart->TextSizePt : 9.0;
        for ($i = 0; $i < $this->LegendItemCount($chart); $i++) {
            $textWidth = max($textWidth, $this->MeasureText($this->LegendItemLabel($chart, $i), $size));
        }
        return $textWidth + $this->ResolveLegendSwatchWidth($chart) + $this->ResolveLegendTextGap($chart) + $size;
    }

    private function LegendItemCount($chart)
    {
        if ($chart === null) {
            return 0;
        }

        if ($chart->Kind === ChartKind::Pie) {
            return $chart->Categories !== null ? count($chart->Categories) : 0;
        }

        return $chart->Series !== null ? count($chart->Series) : 0;
    }

    private function LegendItemLabel($chart, $legendIndex)
    {
        if ($chart === null || $legendIndex < 0) {
            return '';
        }

        if ($chart->Kind === ChartKind::Pie) {
            if ($chart->Categories !== null && $legendIndex < count($chart->Categories)) {
                return $chart->Categories[$legendIndex] !== null ? $chart->Categories[$legendIndex] : '';
            }

            return '';
        }

        $seriesIndex = $this->ResolveLegendSeriesIndex($chart, $legendIndex);
        if ($chart->Series === null || $seriesIndex < 0 || $seriesIndex >= count($chart->Series)) {
            return '';
        }

        return $chart->Series[$seriesIndex]->Name !== null ? $chart->Series[$seriesIndex]->Name : '';
    }

    private function LegendItemColor($chart, $legendIndex)
    {
        if ($chart === null || $legendIndex < 0) {
            return SKColors::Gray();
        }

        if ($chart->Kind === ChartKind::Pie) {
            return $this->context->Colors->ResolveSchemeName('accent' . (($legendIndex % 6) + 1), SKColors::Gray());
        }

        $seriesIndex = $this->ResolveLegendSeriesIndex($chart, $legendIndex);
        if ($chart->Series === null || $seriesIndex < 0 || $seriesIndex >= count($chart->Series)) {
            return SKColors::Gray();
        }

        return $chart->Series[$seriesIndex]->Color;
    }

    private function LegendCenterY($area, $plot, $chart)
    {
        if ($chart !== null && $chart->HasManualLegend) {
            return $area->Top + ($chart->LegendY + $chart->LegendH * 0.5) * $area->Height - 1.67;
        }

        if ($chart !== null && $chart->LegendPosition === 't') {
            return $area->Top + $area->Height * 0.23;
        }

        if ($chart !== null && $chart->Kind === ChartKind::Pie && !$chart->IsOfPie) {
            return $area->Bottom - $area->Height * 0.085;
        }

        $position = 0.68;
        $dateAxisLayout = $this->ResolveDateAxisLayout($chart, $plot->Width);
        if ($chart->HasDateCategoryAxis && abs($dateAxisLayout->RotationDeg) > 0.1) {
            $position = 0.86;
        }
        if ($chart->IsLine3D) {
            $position = 0.76;
        }

        $centerY = $plot->Bottom + ($area->Bottom - $plot->Bottom) * $position;
        if ($chart !== null && $chart->IsBubble) {
            $centerY -= 6.0;
        }
        if ($chart !== null && $chart->Kind === ChartKind::Scatter
            && $chart->CategoryAxisTitle !== null && $chart->CategoryAxisTitle !== '') {
            $centerY = $area->Bottom - 15.7;
        }
        if ($chart !== null && $chart->HasDataTable) {
            $centerY = $area->Top + $area->Height * 0.919;
        } elseif ($chart !== null && $chart->Kind === ChartKind::Column && $chart->IsBar3D) {
            $centerY = $area->Top + $area->Height * 0.93;
        }
        if ($this->IsCompactColumnLegend($chart)) {
            $centerY -= 1.0;
        }
        if ($chart !== null && $chart->IsOfPie && $chart->OfPieType !== null && strcasecmp($chart->OfPieType, 'bar') === 0) {
            $centerY -= 5.0;
        }

        return LineChartDecorationLayout::LegendCenterY($area, $chart, $centerY);
    }

    private function ResolveLegendTextGap($chart)
    {
        if (self::UsesLineScatterLegend($chart)) {
            return 2.0;
        }

        if ($this->IsCompactColumnLegend($chart)) {
            return 2.0;
        }

        if ($chart !== null && $chart->Kind === ChartKind::Pie) {
            return 0.0;
        }

        if ($chart !== null && $chart->Kind === ChartKind::Area) {
            return 0.0;
        }

        return LineChartDecorationLayout::TextGap($chart, 1.0);
    }

    private function ResolveLegendItemGap($chart)
    {
        if ($chart !== null && $chart->Kind === ChartKind::Scatter) {
            if (self::UsesLineScatterLegend($chart)) {
                return 9.5;
            }

            return $chart->HasManualLegend ? 6.9 : 4.6;
        }

        if ($chart !== null && $chart->IsLine3D) {
            return 4.6;
        }

        if ($this->IsCompactColumnLegend($chart)) {
            return 7.45;
        }

        if ($chart !== null && $chart->Kind === ChartKind::Pie) {
            return $chart->IsDoughnut ? 7.1 : 3.6;
        }

        if ($chart !== null && $chart->Kind === ChartKind::Area) {
            return 6.0;
        }

        return LineChartDecorationLayout::ItemGap($chart, 12.6);
    }

    private function ResolveLegendSwatchWidth($chart)
    {
        if (self::UsesLineScatterLegend($chart)) {
            return 19.2;
        }

        if ($chart !== null && $chart->IsDoughnut) {
            return 7.0;
        }

        if ($chart !== null && $chart->HasManualLegend) {
            return 8.25;
        }

        if ($this->IsCompactColumnLegend($chart)) {
            return 5.0;
        }

        return LineChartDecorationLayout::SwatchWidth($chart, 9.0);
    }

    private static function UsesLineScatterLegend($chart)
    {
        if ($chart === null || $chart->Kind !== ChartKind::Scatter || $chart->Series === null) {
            return false;
        }

        for ($i = 0; $i < count($chart->Series); $i++) {
            if ($chart->Series[$i]->HasVisibleLine) {
                return true;
            }
        }

        return false;
    }

    private function ResolveLegendHorizontalOffset($chart)
    {
        if ($chart !== null && $chart->Kind === ChartKind::Scatter) {
            return 1.5;
        }

        if ($chart !== null && $chart->IsDoughnut) {
            return 2.5;
        }

        if ($chart !== null && $chart->IsLine3D) {
            return 3.5;
        }

        if ($this->IsCompactColumnLegend($chart)) {
            return 2.25;
        }

        return LineChartDecorationLayout::HorizontalOffset($chart, 0.0);
    }

    private function ResolveLegendLeft($area, $chart, $totalWidth)
    {
        if ($chart !== null && $chart->HasManualLegend) {
            $center = $area->Left + ($chart->LegendX + $chart->LegendW * 0.5) * $area->Width;
            return $center - $totalWidth * 0.5 + ($chart->Kind === ChartKind::Scatter ? 2.3 : 5.75);
        }

        return $area->MidX - $totalWidth * 0.5 + $this->ResolveLegendHorizontalOffset($chart);
    }

    private function IsCompactColumnLegend($chart)
    {
        return $chart !== null
            && $chart->Kind === ChartKind::Column
            && !$chart->IsPercentStacked
            && !$chart->IsBar3D
            && !$chart->HasManualLegend
            && $chart->Title !== null
            && $chart->LegendPosition !== null;
    }

    private function DrawPie($canvas, $page, $area, $chart)
    {
        $titlePosition = 0.08 - ($chart->IsOfPie ? self::ResolveOfPieWideLayoutFactor($area) * 0.025 : 0.0);
        $this->DrawTitle($canvas, $page, $area, $chart, $area->Top + $area->Height * $titlePosition);
        if (count($chart->Series) === 0) {
            return;
        }

        $values = $chart->Series[0]->Values;
        $sum = 0.0;
        foreach ($values as $v) {
            if ($v !== null) {
                $sum += abs($v);
            }
        }
        if ($sum <= 0.0) {
            return;
        }

        if ($chart->IsOfPie) {
            $this->DrawOfPie($canvas, $page, $area, $chart, $values);
            return;
        }

        if ($chart->IsDoughnut) {
            $doughnutRect = DoughnutChartRenderer::Draw($canvas, $area, $chart);
            $this->DrawPlotAreaBorder($canvas, $page, $area, $this->ResolvePieBorderRect($area, $chart, $doughnutRect), $chart);
            $this->DrawLegend($canvas, $page, $area, $doughnutRect, $chart);
            return;
        }

        if ($chart->IsPie3D) {
            $pointColors = array();
            for ($i = 0; $i < count($values); $i++) {
                $pointColors[] = $this->PiePointColor($chart, $i);
            }

            $ellipse = Pie3DChartRenderer::Draw($canvas, $area, $chart, $values, $pointColors, $sum);
            $this->DrawPlotAreaBorder($canvas, $page, $area, $this->ResolvePieBorderRect($area, $chart, $ellipse), $chart);
            $this->DrawLegend($canvas, $page, $area, $ellipse, $chart);
            return;
        }

        $size = min($area->Width, $area->Height) * 0.635;
        $rect = new SKRect($area->MidX - $size / 2.0, $area->Top + $area->Height * 0.19, $area->MidX + $size / 2.0, $area->Top + $area->Height * 0.19 + $size);
        $start = -90.0;
        for ($i = 0; $i < count($values); $i++) {
            if ($values[$i] === null) {
                continue;
            }
            $sweep = (float)(abs($values[$i]) / $sum * 360.0);

            $fill = new SKPaint();
            $fill->Style = SKPaintStyle::Fill;
            $fill->Color = $this->PiePointColor($chart, $i);
            $fill->IsAntialias = true;

            $path = new SKPath();
            $path->MoveTo($rect->MidX, $rect->MidY);
            $path->ArcTo($rect, $start, $sweep, false);
            $path->Close();
            $canvas->DrawPath($path, $fill);
            if ($i < count($chart->PiePointLineWidthsPt) && $chart->PiePointLineWidthsPt[$i] > 0.0) {
                $outline = new SKPaint();
                $outline->Style = SKPaintStyle::Stroke;
                $outline->Color = $chart->PiePointLineColors[$i];
                $outline->StrokeWidth = (float)$chart->PiePointLineWidthsPt[$i];
                $outline->IsAntialias = true;
                $canvas->DrawPath($path, $outline);
                $outline->Dispose();
            }
            $path->Dispose();
            $fill->Dispose();

            $start += $sweep;
        }

        $this->DrawPlotAreaBorder($canvas, $page, $area, $this->ResolvePieBorderRect($area, $chart, $rect), $chart);
        $this->DrawLegend($canvas, $page, $area, $rect, $chart);
    }

    private function DrawOfPie($canvas, $page, $area, $chart, $values)
    {
        $secondaryIndices = $this->ResolveOfPieSecondaryIndices($chart, $values);
        if (count($secondaryIndices) === 0 || count($secondaryIndices) >= count($values)) {
            $chart->IsOfPie = false;
            $this->DrawPie($canvas, $page, $area, $chart);
            $chart->IsOfPie = true;
            return;
        }

        $plotRect = $this->ResolvePiePlotRect($area, $chart);

        $secondaryFlags = array_fill(0, count($values), false);
        for ($i = 0; $i < count($secondaryIndices); $i++) {
            $secondaryFlags[$secondaryIndices[$i]] = true;
        }

        $primaryValues = array();
        $primaryColors = array();
        $secondaryValues = array();
        $secondaryColors = array();

        $secondarySum = 0.0;
        for ($i = 0; $i < count($values); $i++) {
            if ($values[$i] === null) {
                continue;
            }

            $absolute = abs($values[$i]);
            if ($secondaryFlags[$i]) {
                $secondaryValues[] = $absolute;
                $secondaryColors[] = $this->PiePointColor($chart, $i);
                $secondarySum += $absolute;
            } else {
                $primaryValues[] = $absolute;
                $primaryColors[] = $this->PiePointColor($chart, $i);
            }
        }

        if (count($secondaryValues) === 0 || $secondarySum <= 0.0) {
            $chart->IsOfPie = false;
            $this->DrawPie($canvas, $page, $area, $chart);
            $chart->IsOfPie = true;
            return;
        }

        $primaryValues[] = $secondarySum;
        $primaryColors[] = $this->PiePointColor($chart, count($values));

        $largerChartBonus = $this->ResolveOfPieLargerChartBonus($area);
        $wideLayout = self::ResolveOfPieWideLayoutFactor($area);
        $baseMainSize = min($area->Width, $area->Height) * (0.515 + $largerChartBonus * 0.02);
        $maxMainSize = min($plotRect->Width * (0.64 + $largerChartBonus * 0.03), $plotRect->Height * 0.96);
        $mainSize = min($baseMainSize, $maxMainSize) * (1.22 - 0.12 * $wideLayout);
        $secondSize = $mainSize * (float)max(0.35, min(1.2, $chart->OfPieSecondPieSizePercent / 100.0));
        $gapRatio = (float)max(0.1, min(3.0, $chart->OfPieGapWidthPercent / 100.0));
        $centerGap = 44.465 + 32.643 * $gapRatio + 26.5 * $wideLayout;
        $isBarOfPie = $chart->OfPieType !== null && strcasecmp($chart->OfPieType, 'bar') === 0;
        $wideBarScale = $isBarOfPie ? 1.0 + 0.25 * $wideLayout : 1.0;
        $mainSize *= $wideBarScale;
        $secondSize *= $wideBarScale;
        $centerGap *= $wideBarScale;
        if ($isBarOfPie) {
            $centerGap *= 0.5;
        }

        $secondWidth = $isBarOfPie ? $secondSize * 0.5 : $secondSize;
        $groupWidth = $mainSize + $centerGap + $secondWidth;
        $groupWidthLimit = $plotRect->Width * (0.94 + $largerChartBonus * 0.025 + $wideLayout * 0.055);
        if ($groupWidth > $groupWidthLimit) {
            $scale = $groupWidthLimit / $groupWidth;
            $mainSize *= $scale;
            $secondSize *= $scale;
            $secondWidth *= $scale;
            $centerGap *= $scale;
            $groupWidth = $mainSize + $centerGap + $secondWidth;
        }

        $groupHeight = max($mainSize, $secondSize);
        $left = $plotRect->MidX - $groupWidth / 2.0 - ($isBarOfPie ? 2.0 * $wideLayout : 0.0);
        $top = $plotRect->MidY + $area->Height * 0.018 - $groupHeight / 2.0;
        $mainOffset = $mainSize < $secondSize ? ($secondSize - $mainSize) * 0.5 : 0.0;
        $mainRect = new SKRect($left, $top + $mainOffset, $left + $mainSize, $top + $mainOffset + $mainSize);
        $secondLeft = $mainRect->Right + $centerGap;
        $secondOffset = $secondSize < $mainSize ? ($mainSize - $secondSize) * 0.5 : 0.0;
        $secondRect = new SKRect($secondLeft, $top + $secondOffset, $secondLeft + $secondWidth, $top + $secondOffset + $secondSize);

        $aggregateStart = 0.0;
        $aggregateSweep = 0.0;
        $this->DrawPieSlices($canvas, $mainRect, $primaryValues, $primaryColors, 70.0, $aggregateStart, $aggregateSweep);

        if ($isBarOfPie) {
            $segmentTop = $secondRect->Top;
            for ($i = 0; $i < count($secondaryValues); $i++) {
                $segmentBottom = $i === count($secondaryValues) - 1
                    ? $secondRect->Bottom
                    : $segmentTop + $secondRect->Height * ($secondaryValues[$i] / $secondarySum);

                $fill = new SKPaint();
                $fill->Style = SKPaintStyle::Fill;
                $fill->Color = $secondaryColors[$i];
                $fill->IsAntialias = true;
                $canvas->DrawRect(new SKRect($secondRect->Left, $segmentTop, $secondRect->Right, $segmentBottom), $fill);
                $fill->Dispose();

                $segmentTop = $segmentBottom;
            }
        } else {
            $secondaryStart = 0.0;
            $secondarySweep = 0.0;
            $this->DrawPieSlices($canvas, $secondRect, $secondaryValues, $secondaryColors, $aggregateStart + 145.0, $secondaryStart, $secondarySweep);
        }

        if ($chart->HasOfPieSeparatorLines) {
            $this->DrawOfPieSeparatorLines($canvas, $chart, $mainRect, $aggregateStart, $aggregateSweep, $secondRect);
        }

        $this->DrawPlotAreaBorder($canvas, $page, $area, $this->ResolveOfPieBorderRect($area, $chart), $chart);
        $this->DrawLegend($canvas, $page, $area, $mainRect, $chart);
    }

    private function ResolvePiePlotRect($area, $chart)
    {
        $largerChartBonus = $chart !== null && $chart->IsOfPie ? $this->ResolveOfPieLargerChartBonus($area) : 0.0;
        $left = $area->Left + $area->Width * 0.08;
        $right = $area->Right - $area->Width * 0.08;
        $titleReserve = ($chart !== null ? $chart->Title : null) === null || ($chart !== null ? $chart->Title : null) === '' ? 0.12 : 0.19;
        $legendReserve = $chart !== null && $chart->LegendPosition !== null ? 0.22 : 0.10;
        if ($largerChartBonus > 0.0) {
            $titleReserve -= 0.025 * $largerChartBonus;
            $legendReserve -= 0.03 * $largerChartBonus;
        }

        $top = $area->Top + $area->Height * $titleReserve;
        $bottom = $area->Bottom - $area->Height * $legendReserve;
        $rect = new SKRect($left, $top, $right, $bottom);

        if ($largerChartBonus > 0.0) {
            $rect = self::ExpandRectAboutCenter($rect, 1.0 + 0.10 * $largerChartBonus, 1.0 + 0.20 * $largerChartBonus);
            $rect = self::InsetAndClamp($area, $rect->Left, $rect->Top, $rect->Right, $rect->Bottom);
        }

        return $rect;
    }

    private function ResolvePieBorderRect($area, $chart, $contentRect)
    {
        $horizontalPadding = $area->Width * 0.035;
        $verticalPadding = $area->Height * 0.03;
        $left = $contentRect->Left - $horizontalPadding;
        $right = $contentRect->Right + $horizontalPadding;
        $top = $contentRect->Top - $verticalPadding;
        $bottom = $contentRect->Bottom + $verticalPadding;

        $title = $chart !== null ? $chart->Title : null;
        $titleBottom = $area->Top + (($title === null || $title === '') ? $area->Height * 0.08 : $area->Height * 0.18);
        if ($top < $titleBottom) {
            $top = $titleBottom;
        }

        $legendTop = $area->Bottom - ($chart !== null && $chart->LegendPosition !== null ? $area->Height * 0.24 : $area->Height * 0.08);
        if ($bottom > $legendTop) {
            $bottom = $legendTop;
        }

        return self::InsetAndClamp($area, $left, $top, $right, $bottom);
    }

    private function ResolveOfPieBorderRect($area, $chart)
    {
        $plotRect = $this->ResolvePiePlotRect($area, $chart);
        $largerChartBonus = $this->ResolveOfPieLargerChartBonus($area);
        $wideLayout = self::ResolveOfPieWideLayoutFactor($area);
        $borderRect = self::ExpandRectAboutCenter($plotRect, 1.0 + 0.116 * $largerChartBonus, 1.29);

        $title = $chart !== null ? $chart->Title : null;
        $topLimit = $area->Top + (($title === null || $title === '') ? $area->Height * 0.06 : $area->Height * (0.12 + 0.013 * $wideLayout));
        if ($borderRect->Top < $topLimit) {
            $borderRect->Top = $topLimit;
        }

        $bottomLimit = $area->Bottom - ($chart !== null && $chart->LegendPosition !== null ? $area->Height * (0.18 - 0.054 * $wideLayout) : $area->Height * 0.05);
        if ($borderRect->Bottom > $bottomLimit) {
            $borderRect->Bottom = $bottomLimit;
        }

        $sideInset = $area->Width * 0.026 * $wideLayout;
        return self::InsetAndClamp($area, max($borderRect->Left, $area->Left + $sideInset), $borderRect->Top,
            min($borderRect->Right, $area->Right - $sideInset), $borderRect->Bottom);
    }

    private function DrawPieSlices($canvas, $rect, $values, $colors, $initialStart, &$lastStart, &$lastSweep)
    {
        $lastStart = $initialStart;
        $lastSweep = 0.0;
        $sum = 0.0;
        for ($i = 0; $i < count($values); $i++) {
            $sum += $values[$i];
        }

        if ($sum <= 0.0) {
            return;
        }

        $start = $initialStart;
        for ($i = 0; $i < count($values); $i++) {
            $sweep = (float)($values[$i] / $sum * 360.0);

            $fill = new SKPaint();
            $fill->Style = SKPaintStyle::Fill;
            $fill->Color = $i < count($colors) ? $colors[$i] : SKColors::Gray();
            $fill->IsAntialias = true;

            $path = new SKPath();
            $path->MoveTo($rect->MidX, $rect->MidY);
            $path->ArcTo($rect, $start, $sweep, false);
            $path->Close();
            $canvas->DrawPath($path, $fill);

            $path->Dispose();
            $fill->Dispose();

            $lastStart = $start;
            $lastSweep = $sweep;
            $start += $sweep;
        }
    }

    private function DrawOfPieSeparatorLines($canvas, $chart, $mainRect, $aggregateStart, $aggregateSweep, $secondRect)
    {
        $topMain = null;
        $topSecond = null;
        $bottomMain = null;
        $bottomSecond = null;
        if ($chart->OfPieType !== null && strcasecmp($chart->OfPieType, 'bar') === 0) {
            $topSecond = new SKPoint($secondRect->Left, $secondRect->Top);
            $bottomSecond = new SKPoint($secondRect->Left, $secondRect->Bottom);
        } elseif (!self::TryResolveOuterTangents($mainRect, $secondRect, $topMain, $topSecond, $bottomMain, $bottomSecond)) {
            $firstAngle = $aggregateStart;
            $secondAngle = $aggregateStart + $aggregateSweep;
            $topMain = self::PieArcPoint($mainRect, $firstAngle);
            $bottomMain = self::PieArcPoint($mainRect, $secondAngle);
            $topSecond = new SKPoint($secondRect->Left, $secondRect->MidY - $secondRect->Height * 0.22);
            $bottomSecond = new SKPoint($secondRect->Left, $secondRect->MidY + $secondRect->Height * 0.22);

            if ($topMain->Y > $bottomMain->Y) {
                $swap = $topMain;
                $topMain = $bottomMain;
                $bottomMain = $swap;
            }
        }

        $topMain = self::PieArcPoint($mainRect, $aggregateStart);
        $bottomMain = self::PieArcPoint($mainRect, $aggregateStart + $aggregateSweep);

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->OfPieSeparatorLineColor;
        $paint->StrokeWidth = (float)max(0.5, $chart->OfPieSeparatorLineWidthPt);
        $paint->IsAntialias = true;

        $canvas->DrawLine($topMain, $topSecond, $paint);
        $canvas->DrawLine($bottomMain, $bottomSecond, $paint);

        $paint->Dispose();
    }

    private static function TryResolveOuterTangents($leftCircle, $rightCircle, &$topLeft, &$topRight, &$bottomLeft, &$bottomRight)
    {
        $topLeft = new SKPoint(0.0, 0.0);
        $topRight = new SKPoint(0.0, 0.0);
        $bottomLeft = new SKPoint(0.0, 0.0);
        $bottomRight = new SKPoint(0.0, 0.0);

        $x1 = $leftCircle->MidX;
        $y1 = $leftCircle->MidY;
        $r1 = $leftCircle->Width * 0.5;
        $x2 = $rightCircle->MidX;
        $y2 = $rightCircle->MidY;
        $r2 = $rightCircle->Width * 0.5;

        $dx = $x2 - $x1;
        $dy = $y2 - $y1;
        $distanceSquared = $dx * $dx + $dy * $dy;
        if ($distanceSquared <= 0.0001) {
            return false;
        }

        $radiusDelta = $r1 - $r2;
        $hSquared = $distanceSquared - $radiusDelta * $radiusDelta;
        if ($hSquared <= 0.0001) {
            return false;
        }

        $inv = 1.0 / $distanceSquared;
        $h = sqrt($hSquared);
        $nxTop = ($dx * $radiusDelta - $dy * $h) * $inv;
        $nyTop = ($dy * $radiusDelta + $dx * $h) * $inv;
        $nxBottom = ($dx * $radiusDelta + $dy * $h) * $inv;
        $nyBottom = ($dy * $radiusDelta - $dx * $h) * $inv;

        $topLeft = new SKPoint($x1 + $r1 * $nxTop, $y1 + $r1 * $nyTop);
        $topRight = new SKPoint($x2 + $r2 * $nxTop, $y2 + $r2 * $nyTop);
        $bottomLeft = new SKPoint($x1 + $r1 * $nxBottom, $y1 + $r1 * $nyBottom);
        $bottomRight = new SKPoint($x2 + $r2 * $nxBottom, $y2 + $r2 * $nyBottom);

        if ($topLeft->Y > $bottomLeft->Y) {
            $swapLeft = $topLeft;
            $topLeft = $bottomLeft;
            $bottomLeft = $swapLeft;

            $swapRight = $topRight;
            $topRight = $bottomRight;
            $bottomRight = $swapRight;
        }

        return true;
    }

    private static function PieArcPoint($rect, $angleDeg)
    {
        $radians = M_PI / 180.0 * $angleDeg;
        return new SKPoint(
            $rect->MidX + cos($radians) * $rect->Width * 0.5,
            $rect->MidY + sin($radians) * $rect->Height * 0.5
        );
    }

    private function ResolveOfPieSecondaryIndices($chart, $values)
    {
        $result = array();
        if ($chart === null || $values === null) {
            return $result;
        }

        $validCount = 0;
        for ($i = 0; $i < count($values); $i++) {
            if ($values[$i] !== null) {
                $validCount++;
            }
        }

        if ($validCount <= 1) {
            return $result;
        }

        $splitType = $chart->OfPieSplitType;
        if ($splitType === null || $splitType === '' || $splitType === 'auto' || $splitType === 'pos') {
            $count = $chart->OfPieSplitPosition !== null ? (int)round($chart->OfPieSplitPosition) : 2;
            if ($count < 1) {
                $count = 1;
            }

            if ($count >= $validCount) {
                $count = $validCount - 1;
            }

            for ($i = count($values) - 1; $i >= 0 && count($result) < $count; $i--) {
                if ($values[$i] !== null) {
                    array_unshift($result, $i);
                }
            }

            return $result;
        }

        if ($splitType === 'val' && $chart->OfPieSplitPosition !== null) {
            for ($i = 0; $i < count($values); $i++) {
                if ($values[$i] !== null && abs($values[$i]) <= $chart->OfPieSplitPosition) {
                    $result[] = $i;
                }
            }

            if (count($result) > 0 && count($result) < $validCount) {
                return $result;
            }
        }

        if ($splitType === 'percent' && $chart->OfPieSplitPosition !== null) {
            $sum = 0.0;
            for ($i = 0; $i < count($values); $i++) {
                if ($values[$i] !== null) {
                    $sum += abs($values[$i]);
                }
            }

            if ($sum > 0.0) {
                for ($i = 0; $i < count($values); $i++) {
                    if ($values[$i] !== null && abs($values[$i]) / $sum * 100.0 <= $chart->OfPieSplitPosition) {
                        $result[] = $i;
                    }
                }

                if (count($result) > 0 && count($result) < $validCount) {
                    return $result;
                }
            }
        }

        for ($i = count($values) - 1; $i >= 0 && count($result) < min(3, $validCount - 1); $i--) {
            if ($values[$i] !== null) {
                array_unshift($result, $i);
            }
        }

        return $result;
    }

    private function PiePointColor($chart, $pointIndex)
    {
        if ($chart !== null && $chart->PiePointColors !== null && $pointIndex >= 0 && $pointIndex < count($chart->PiePointColors)) {
            return $chart->PiePointColors[$pointIndex];
        }

        return $this->context->Colors->ResolveSchemeName('accent' . (($pointIndex % 6) + 1), SKColors::Gray());
    }

    private static function MaxPointCount($chart)
    {
        $n = $chart->Categories !== null ? count($chart->Categories) : 0;
        foreach ($chart->Series as $s) {
            $n = max($n, count($s->Values));
        }
        return $n;
    }

    private function DrawDateCategoryAxis($canvas, $page, $area, $plot, $chart)
    {
        if ($chart->Kind === ChartKind::Column && $chart->CrossBetween) {
            $this->DrawColumnDateCategoryAxis($canvas, $page, $area, $plot, $chart);
            return;
        }

        $plotMinDate = 0.0;
        $plotMaxDate = 0.0;
        if (!ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $this->dateSystem, $plotMinDate, $plotMaxDate)) {
            return;
        }

        $minDate = 0.0;
        $maxDate = 0.0;
        if (!ChartDateAxisMath::TryGetDateAxisRange($chart, $minDate, $maxDate)) {
            return;
        }

        $firstTick = ChartDateAxisMath::StartOfMonth($minDate, $this->dateSystem);
        $lastTick = ChartDateAxisMath::StartOfMonth($maxDate, $this->dateSystem);
        $layout = $this->ResolveDateAxisLayout($chart, $plot->Width);
        $boundaryGridlines = LineChartCategoryAxisRenderer::DrawBoundaryGridlines($canvas, $plot, $chart, $this->dateSystem);
        $visibleSlice = $this->ResolveVisibleChartSlice($page, $area);
        $current = $firstTick;
        $rotatedClip = new SKRect(
            max($plot->Left - 36.0, $visibleSlice->Left),
            $plot->Bottom + 2.0,
            min($plot->Right + 18.0, $visibleSlice->Right),
            $visibleSlice->Bottom
        );
        while ($current <= $lastTick + 0.1) {
            $x = ChartDateAxisMath::DateCategoryX($current, $plotMinDate, $plotMaxDate, $plot, $chart, $this->dateSystem);
            if ($chart->HasCategoryGridlines && !$boundaryGridlines) {
                if (!self::IsInteriorVerticalGridline($x, $plot)) {
                    $current = ChartDateAxisMath::AddMonthsSerial($current, $layout->MonthStep, $this->dateSystem);
                    continue;
                }

                $grid = new SKPaint();
                $grid->Style = SKPaintStyle::Stroke;
                $grid->Color = $chart->CategoryGridlineColor;
                $grid->StrokeWidth = (float)max(0.5, $chart->CategoryGridlineWidthPt);
                $grid->IsAntialias = false;
                $canvas->DrawLine($x, $plot->Top, $x, $plot->Bottom, $grid);
                $grid->Dispose();
            }

            if ($this->suppressCategoryAxisLabels) {
                $current = ChartDateAxisMath::AddMonthsSerial($current, $layout->MonthStep, $this->dateSystem);
                continue;
            }

            $label = $this->FormatValue($current, ($chart->CategoryFormatCode === null || $chart->CategoryFormatCode === '') ? 'mmm-yy' : $chart->CategoryFormatCode);
            if (abs($layout->RotationDeg) > 0.1) {
                $allowTrailingClipOnLastSlice = $this->IsLastHorizontalChartSlice($page, $area);
                $allowLeadingClipOnContinuationSlice = $area->Left < $visibleSlice->Left - 0.1;
                $leftCull = $allowLeadingClipOnContinuationSlice ? $visibleSlice->Left - 72.0 : $visibleSlice->Left - 10.0;
                if ($x < $leftCull || (!$allowTrailingClipOnLastSlice && $x > $visibleSlice->Right + 24.0)) {
                    $current = ChartDateAxisMath::AddMonthsSerial($current, $layout->MonthStep, $this->dateSystem);
                    continue;
                }

                $automaticDiagonal = LineChartAutomaticLayout::Applies($chart) && abs($layout->RotationDeg + 45.0) < 0.1;
                $labelX = $automaticDiagonal ? $x - $this->LabelSizePt * 0.5 : $x + 2.0;
                $this->DrawRotatedText($canvas, $page, $rotatedClip, $label, $labelX, $plot->Bottom + ($automaticDiagonal ? 4.0 : 6.0), $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Right, VAlign::Top, $layout->RotationDeg);
            } else {
                $this->DrawText($canvas, $page, $area, $label, $x, $plot->Bottom + 4.0, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Center, VAlign::Top);
            }
            $current = ChartDateAxisMath::AddMonthsSerial($current, $layout->MonthStep, $this->dateSystem);
        }
    }

    private function DrawColumnDateCategoryAxis($canvas, $page, $area, $plot, $chart)
    {
        $count = max(count($chart->Categories), count($chart->CategoryValues));
        if ($count <= 0) {
            return;
        }

        if ($chart->HasCategoryGridlines) {
            $grid = new SKPaint();
            $grid->Style = SKPaintStyle::Stroke;
            $grid->Color = $chart->CategoryGridlineColor;
            $grid->StrokeWidth = (float)max(0.5, $chart->CategoryGridlineWidthPt);
            $grid->IsAntialias = false;
            for ($i = 0; $i <= $count; $i++) {
                $x = self::ColumnDateCategoryBoundaryX($i, $count, $plot);
                $canvas->DrawLine($x, $plot->Top, $x, $plot->Bottom, $grid);
            }
            $grid->Dispose();
        }

        if ($this->suppressCategoryAxisLabels) {
            return;
        }

        for ($i = 0; $i < $count; $i++) {
            $label = $this->ResolveCategoryLabel($chart, $i);
            $x = self::CategoryX($i, $count, $plot, true);
            $this->DrawText($canvas, $page, $area, $label, $x, $plot->Bottom + 4.0, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Center, VAlign::Top);
        }
    }

    private static function ColumnDateCategoryBoundaryX($index, $count, $plot)
    {
        if ($count <= 0) {
            return $plot->Left;
        }

        return $plot->Left + (float)$index / $count * $plot->Width;
    }

    private function DrawBarValueAxis($canvas, $page, $area, $plot, $chart, $min, $max, $major)
    {
        if ($max <= $min || $major <= 0.0) {
            return;
        }

        $major = self::ResolveRenderedMajorUnit($chart, $plot, $major);

        $grid = new SKPaint();
        $grid->Style = SKPaintStyle::Stroke;
        $grid->Color = self::GridColor();
        $grid->StrokeWidth = 1.0;
        $grid->IsAntialias = false;

        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $x = self::ValueToX($value, $min, $max, $plot);

            if ($chart->HasValueGridlines) {
                $canvas->DrawLine($x, $plot->Top, $x, $plot->Bottom, $grid);
            }

            $label = $this->FormatValue($value, $chart->ValueFormatCode);
            $width = $this->MeasureText($label, $this->LabelSizePt);
            $clip = new SKRect($x - $width * 0.5 - 2.0, $plot->Bottom + 2.0, $x + $width * 0.5 + 2.0, $plot->Bottom + $this->LabelSizePt * 2.2);
            $this->DrawText($canvas, $page, $clip, $label, $x, $plot->Bottom + 4.0, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Center, VAlign::Top);
        }

        $grid->Dispose();
    }

    private function DrawBarCategoryAxis($canvas, $page, $area, $plot, $chart)
    {
        $axis = new SKPaint();
        $axis->Style = SKPaintStyle::Stroke;
        $axis->Color = self::GridColor();
        $axis->StrokeWidth = 1.0;
        $axis->IsAntialias = false;
        $canvas->DrawLine($plot->Left, $plot->Top, $plot->Left, $plot->Bottom, $axis);
        $axis->Dispose();

        if ($this->suppressCategoryAxisLabels) {
            return;
        }

        $count = count($chart->Categories);
        if ($count === 0) {
            $count = count($chart->CategoryValues);
        }

        if ($count === 0) {
            return;
        }

        for ($i = 0; $i < $count; $i++) {
            if ($chart->HasCategoryGridlines) {
                $gridlineY = $this->CategoryGridlineY($i, $count, $plot, $chart->CrossBetween, $this->ShouldReverseBarCategoryAxis($chart));

                $grid = new SKPaint();
                $grid->Style = SKPaintStyle::Stroke;
                $grid->Color = $chart->CategoryGridlineColor;
                $grid->StrokeWidth = (float)max(0.5, $chart->CategoryGridlineWidthPt);
                $grid->IsAntialias = false;
                $canvas->DrawLine($plot->Left, $gridlineY, $plot->Right, $gridlineY, $grid);
                $grid->Dispose();
            }

            $y = $this->CategoryY($i, $count, $plot, $chart->CrossBetween, $this->ShouldReverseBarCategoryAxis($chart));
            $label = $this->ResolveCategoryLabel($chart, $i);
            if ($label === null || $label === '') {
                continue;
            }

            $width = $this->MeasureText($label, $this->LabelSizePt);
            $clip = new SKRect($plot->Left - $width - 8.0, $y - $this->LabelSizePt, $plot->Left - 2.0, $y + $this->LabelSizePt);
            $this->DrawText($canvas, $page, $clip, $label, $plot->Left - 5.0, $y, $this->LabelSizePt, false, $chart->AxisLabelColor, TextAlign::Right, VAlign::Middle);
        }
    }

    private function PointX($chart, $index, $count, $plot)
    {
        if ($chart !== null && $chart->HasDateCategoryAxis) {
            $plotMinDate = 0.0;
            $plotMaxDate = 0.0;
            if (ChartDateAxisMath::TryGetDateAxisPlotRange($chart, $this->dateSystem, $plotMinDate, $plotMaxDate)
                && $index >= 0
                && $index < count($chart->CategoryValues)
                && $chart->CategoryValues[$index] !== null) {
                $serial = ChartDateAxisMath::NormalizePointSerial($chart, $chart->CategoryValues[$index], $this->dateSystem);
                return ChartDateAxisMath::DateCategoryX($serial, $plotMinDate, $plotMaxDate, $plot, $chart, $this->dateSystem);
            }
        }

        return self::CategoryX($index, $count, $plot, $chart !== null && $chart->CrossBetween);
    }

    private function CategoryY($index, $count, $plot, $crossBetween, $reverse = false)
    {
        if ($count <= 0) {
            return $plot->MidY;
        }

        $renderIndex = $reverse ? $count - 1 - $index : $index;

        if ($crossBetween) {
            return $plot->Top + ($renderIndex + 0.5) / $count * $plot->Height;
        }

        if ($count === 1) {
            return $plot->MidY;
        }

        return $plot->Top + (float)$renderIndex / ($count - 1) * $plot->Height;
    }

    private function CategoryGridlineY($index, $count, $plot, $crossBetween, $reverse)
    {
        if ($count <= 1) {
            return $plot->MidY;
        }

        if ($crossBetween) {
            $renderIndex = $reverse ? $count - 1 - $index : $index;
            if ($renderIndex >= $count - 1) {
                $renderIndex = $count - 2;
            }

            return $plot->Top + ($renderIndex + 1.0) / $count * $plot->Height;
        }

        return $this->CategoryY($index, $count, $plot, false, $reverse);
    }

    private function ResolveBarCategoryIndex($chart, $index, $count)
    {
        if (!$this->ShouldReverseBarCategoryAxis($chart)) {
            return $index;
        }

        return $count - 1 - $index;
    }

    private function ShouldReverseBarCategoryAxis($chart)
    {
        if ($chart === null || $chart->Kind !== ChartKind::Bar) {
            return false;
        }

        return $chart->CategoryAxisMinMax;
    }

    private function ResolveCategoryLabel($chart, $index)
    {
        if ($chart === null || $index < 0) {
            return '';
        }

        if ($index < count($chart->Categories) && $chart->Categories[$index] !== null && $chart->Categories[$index] !== '') {
            return $chart->Categories[$index];
        }

        if ($index < count($chart->CategoryValues) && $chart->CategoryValues[$index] !== null) {
            return $this->FormatValue($chart->CategoryValues[$index], ($chart->CategoryFormatCode === null || $chart->CategoryFormatCode === '') ? 'mmm-yy' : $chart->CategoryFormatCode);
        }

        return '';
    }

    private static function CategoryX($index, $count, $plot, $crossBetween)
    {
        if ($count <= 0) {
            return $plot->MidX;
        }

        if ($crossBetween) {
            // Points sit at the center of each category band (Excel's default), inset from the
            // plot edges by half a band.
            return $plot->Left + ($index + 0.5) / $count * $plot->Width;
        }

        if ($count === 1) {
            return $plot->MidX;
        }

        // Points sit on the axis ticks: first at the left edge, last at the right edge.
        return $plot->Left + (float)$index / ($count - 1) * $plot->Width;
    }

    private static function ValueToY($value, $min, $max, $plot)
    {
        $t = ($value - $min) / ($max - $min);
        return $plot->Bottom - (float)$t * $plot->Height;
    }

    private static function ValueToX($value, $min, $max, $plot)
    {
        $t = ($value - $min) / ($max - $min);
        return $plot->Left + (float)$t * $plot->Width;
    }

    private static function FirstValueIndex($values)
    {
        for ($i = 0; $i < count($values); $i++) {
            if ($values[$i] !== null) {
                return $i;
            }
        }
        return 0;
    }

    private static function LastValueIndex($values)
    {
        for ($i = count($values) - 1; $i >= 0; $i--) {
            if ($values[$i] !== null) {
                return $i;
            }
        }
        return 0;
    }

    private function ResolveVisibleChartSlice($page, $area)
    {
        if ($page === null || $page->Sheet === null) {
            return $area;
        }

        $contentWidth = (float)($page->Sheet->ColumnStartPt[$page->EndColumn + 1] - $page->Sheet->ColumnStartPt[$page->StartColumn]);
        $contentHeight = (float)($page->Sheet->RowStartPt[$page->EndRow + 1] - $page->Sheet->RowStartPt[$page->StartRow]);
        $left = max($area->Left, 0.0);
        $top = max($area->Top, 0.0);
        $right = min($area->Right, $contentWidth);
        $bottom = min($area->Bottom, $contentHeight);
        return new SKRect($left, $top, max($left, $right), max($top, $bottom));
    }

    private static function ComputeValueAxis($chart, &$min, &$max, &$major)
    {
        $dmin = 0.0;
        $dmax = 0.0;
        ChartStackingMath::ComputeDataRange($chart, $dmin, $dmax);

        if ($dmin === PHP_FLOAT_MAX) {
            $min = 0.0;
            $max = 1.0;
            $major = 0.2;
            return;
        }

        if ($chart->IsBubble) {
            BubbleChartMath::ResolveYRange($chart, $dmin, $dmax, $min, $max, $major);
            return;
        }

        if ($chart !== null && $chart->IsPercentStacked) {
            $min = $chart->AxisMin !== null ? $chart->AxisMin : 0.0;
            $max = $chart->AxisMax !== null ? $chart->AxisMax : 1.0;
            $automaticMajor = $chart->Kind === ChartKind::Column && $chart->IsBar3D
                && $chart->BarShape === 'cylinder'
                ? 0.5
                : 0.2;
            $major = $chart->MajorUnit !== null ? $chart->MajorUnit : $automaticMajor;
            if ($major <= 0.0) {
                $major = $automaticMajor;
            }

            if ($max <= $min) {
                $max = $min + $major;
            }

            return;
        }

        $min = $chart->AxisMin !== null ? $chart->AxisMin : ($dmin >= 0.0 ? 0.0 : -self::NiceNumber(-$dmin));
        $compactArea3D = $chart !== null
            && $chart->IsArea3D
            && count($chart->Series) > max(count($chart->Categories), count($chart->CategoryValues));
        $labeledColumn3D = false;
        if ($chart !== null && $chart->Kind === ChartKind::Column && $chart->IsBar3D) {
            for ($seriesIndex = 0; $seriesIndex < count($chart->Series); $seriesIndex++) {
                if ($chart->Series[$seriesIndex]->ShowValueDataLabels) {
                    $labeledColumn3D = true;
                    break;
                }
            }
        }
        if ($chart !== null && $chart->IsLine3D) {
            $automaticMajorIntervals = 4.0;
        } elseif ($chart !== null && $chart->Kind === ChartKind::Scatter
            && $chart->ValueFormatCode !== null && $chart->ValueFormatCode !== '' && strpos($chart->ValueFormatCode, '%') !== false) {
            $automaticMajorIntervals = 6.0;
        } elseif ($compactArea3D && $chart->IsStacked) {
            $automaticMajorIntervals = $chart->PerspectiveDeg >= 100.0 ? 2.0 : 3.0;
        } elseif ($chart !== null && $chart->Kind === ChartKind::Column && $chart->IsBar3D) {
            $automaticMajorIntervals = $labeledColumn3D ? 7.0 : 4.0;
        } else {
            $automaticMajorIntervals = 8.0;
        }
        $major = $chart->MajorUnit !== null ? $chart->MajorUnit : self::NiceNumber(($dmax - $min) / $automaticMajorIntervals);
        if ($major <= 0.0) {
            $major = 1.0;
        }

        if ($chart->AxisMax !== null) {
            $max = $chart->AxisMax;
        } elseif ($chart->IsStacked && !$chart->IsBar3D && $min == 0.0 && $dmax > 0.0
            && ($chart->Kind === ChartKind::Column || $chart->Kind === ChartKind::Bar)) {
            // Positive 2-D stacks reserve 5% above the total before rounding to a tick.
            // Excel probes: 571 -> 600, 572 -> 700 (major unit 100).
            $max = ceil($dmax * 1.05 / $major) * $major;
        } else {
            $max = ceil(($dmax - $min) / $major) * $major + $min;
            // Radar charts use the rounded outer ring directly instead of adding
            // an extra headroom band above the data maximum.
            $denseBar3DGrid = $chart !== null && $chart->IsBar3D && $chart->HasMinorValueGridlines;
            $standardDateBar3D = $chart !== null && $chart->IsBar3D
                && $chart->HasDateCategoryAxis && $chart->PerspectiveDeg < 100.0;
            if (($chart === null || $chart->Kind !== ChartKind::Radar) && !$compactArea3D && !$denseBar3DGrid
                && !$standardDateBar3D
                && $max - $dmax < 0.15 * $major) {
                $max += $major;
            }
        }

        if ($max <= $min) {
            $max = $min + $major;
        }
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

    private function FormatValue($value, $formatCode)
    {
        $text = ChartXmlParser::FormatLabel($value, $formatCode, $this->dateSystem, $this->chartCulture);
        return $text !== null ? rtrim($text) : '';
    }

    private function MeasureMaxValueAxisLabelWidth($chart)
    {
        $min = 0.0;
        $max = 0.0;
        $major = 0.0;
        self::ComputeValueAxis($chart, $min, $max, $major);
        if ($max <= $min || $major <= 0.0) {
            return 0.0;
        }

        $maxWidth = 0.0;
        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = $min + $i * $major;
            $label = $this->FormatValue($value / $chart->ValueDisplayUnitDivisor, $chart->ValueFormatCode);
            $width = $this->MeasureText($label, $this->LabelSizePt);
            if ($width > $maxWidth) {
                $maxWidth = $width;
            }
        }

        return $maxWidth;
    }

    private static function InsetAndClamp($area, $left, $top, $right, $bottom)
    {
        if ($left < $area->Left) {
            $left = $area->Left;
        }

        if ($top < $area->Top) {
            $top = $area->Top;
        }

        if ($right > $area->Right) {
            $right = $area->Right;
        }

        if ($bottom > $area->Bottom) {
            $bottom = $area->Bottom;
        }

        if ($right <= $left) {
            $left = $area->Left;
            $right = $area->Right;
        }

        if ($bottom <= $top) {
            $top = $area->Top;
            $bottom = $area->Bottom;
        }

        return new SKRect($left, $top, $right, $bottom);
    }

    private function ResolveDateAxisLayout($chart, $availableWidth)
    {
        return ChartDateAxisLayout::Resolve($chart, $this->dateSystem, $availableWidth);
    }

    private function DrawPlotAreaBorder($canvas, $page, $area, $plot, $chart)
    {
        if ($canvas === null || $chart === null || !$chart->HasPlotAreaBorder) {
            return;
        }

        $strokeWidth = (float)max(0.5, $chart->PlotAreaBorderWidthPt);
        if ($chart->Kind === ChartKind::Pie) {
            $strokeWidth = max(0.5, $strokeWidth * 0.75);
        }

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $chart->PlotAreaBorderColor;
        $paint->StrokeWidth = $strokeWidth;
        $paint->IsAntialias = false;

        if ($chart->Kind === ChartKind::Scatter) {
            $canvas->DrawRect($plot, $paint);
            $paint->Dispose();
            return;
        }

        if ($chart->Kind === ChartKind::Pie) {
            $canvas->DrawLine($plot->Left, $plot->Top, $plot->Right, $plot->Top, $paint);
            $canvas->DrawLine($plot->Left, $plot->Bottom, $plot->Right, $plot->Bottom, $paint);
            $canvas->DrawLine($plot->Left, $plot->Top, $plot->Left, $plot->Bottom, $paint);
            $canvas->DrawLine($plot->Right, $plot->Top, $plot->Right, $plot->Bottom, $paint);
            $paint->Dispose();
            return;
        }

        $drawRightEdge = $this->IsLastHorizontalChartSlice($page, $area);
        $verticalStrokeWidth = $chart->Kind === ChartKind::Line ? $paint->StrokeWidth : max(0.5, $paint->StrokeWidth * 0.7);
        $canvas->DrawLine($plot->Left, $plot->Top, $plot->Right, $plot->Top, $paint);
        $paint->StrokeWidth = $verticalStrokeWidth;
        $canvas->DrawLine($plot->Left, $plot->Top, $plot->Left, $plot->Bottom, $paint);
        if ($drawRightEdge) {
            $canvas->DrawLine($plot->Right, $plot->Top, $plot->Right, $plot->Bottom, $paint);
        }

        $paint->Dispose();
    }

    private static function UnionRect($a, $b)
    {
        return new SKRect(
            min($a->Left, $b->Left),
            min($a->Top, $b->Top),
            max($a->Right, $b->Right),
            max($a->Bottom, $b->Bottom)
        );
    }

    private static function ExpandRectAboutCenter($rect, $scaleX, $scaleY)
    {
        $halfWidth = $rect->Width * $scaleX * 0.5;
        $halfHeight = $rect->Height * $scaleY * 0.5;
        return new SKRect(
            $rect->MidX - $halfWidth,
            $rect->MidY - $halfHeight,
            $rect->MidX + $halfWidth,
            $rect->MidY + $halfHeight
        );
    }

    private function ResolveOfPieLargerChartBonus($area)
    {
        $minDimension = min($area->Width, $area->Height);
        if ($minDimension <= 170.0) {
            return 0.0;
        }

        if ($minDimension >= 230.0) {
            return 1.0;
        }

        return ($minDimension - 170.0) / 60.0;
    }

    private static function ResolveOfPieWideLayoutFactor($area)
    {
        return max(0.0, min(1.0, (min($area->Width, $area->Height) - 200.0) / 70.0));
    }

    private static function IsInteriorVerticalGridline($x, $plot)
    {
        return $x > $plot->Left + 0.5 && $x < $plot->Right - 0.5;
    }

    private function IsLastHorizontalChartSlice($page, $area)
    {
        if ($page === null || $page->Sheet === null) {
            return true;
        }

        $visibleSlice = $this->ResolveVisibleChartSlice($page, $area);
        return $area->Right <= $visibleSlice->Right + 0.1;
    }

    private static function IntersectRect($a, $b)
    {
        $left = max($a->Left, $b->Left);
        $top = max($a->Top, $b->Top);
        $right = min($a->Right, $b->Right);
        $bottom = min($a->Bottom, $b->Bottom);
        if ($right < $left) {
            $right = $left;
        }

        if ($bottom < $top) {
            $bottom = $top;
        }

        return new SKRect($left, $top, $right, $bottom);
    }

    private static function ResolveRenderedMajorUnit($chart, $plot, $major)
    {
        return ChartValueAxisIntervals::ResolveMajor($chart, $plot, $major);
    }

    private static function ResolveRenderedMinorUnit($chart, $major)
    {
        if ($chart === null) {
            return 0.0;
        }

        if ($chart->MinorUnit !== null && $chart->MinorUnit > 0.0) {
            return $chart->MinorUnit;
        }

        if ($major <= 0.0) {
            return 0.0;
        }

        return $major / 5.0;
    }

    private function ResolveValueGridlineRight($plot, $chart)
    {
        if ($chart !== null
            && $chart->Kind === ChartKind::Line
            && !$chart->IsStacked
            && !$chart->IsPercentStacked
            && $chart->HasDateCategoryAxis
            && $chart->LegendPosition !== null) {
            return LineChartAutomaticLayout::Applies($chart) && !$chart->HasManualPlot
                ? $plot->Right : $plot->Right + self::StandardDateAxisGridlineRightExtensionPt;
        }

        if ($chart !== null
            && $chart->Kind === ChartKind::Line
            && $chart->IsStacked
            && !$chart->IsPercentStacked
            && $chart->HasDateCategoryAxis
            && $chart->LegendPosition !== null) {
            return LineChartAutomaticLayout::Applies($chart) && !$chart->HasManualPlot
                ? $plot->Right : $plot->Right + self::StackedDateAxisGridlineRightExtensionPt;
        }

        return $plot->Right;
    }

    private function MeasureText($text, $sizePt, $bold = false)
    {
        if (is_object($sizePt)) {
            $font = $sizePt;
        } else {
            $font = $this->ChartFont((float)$sizePt, $bold);
        }

        $fontContext = $this->context->GetFontContext($font);
        return $fontContext->Measure($text !== null ? $text : '');
    }

    private function DrawText($canvas, $page, $clipRect, $text, $x, $y, $sizePt, $bold, $color, $hAlign, $vAlign, $outlineColor = null, $italic = false)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = $this->ChartFont($sizePt, $bold, $italic);
        $fontContext = $this->context->GetFontContext($fontValue);
        $paint = $this->TextPaint($fontValue, $color);

        $width = $fontContext->Measure($text);
        $drawX = $hAlign === TextAlign::Left ? $x : ($hAlign === TextAlign::Center ? $x - $width / 2.0 : $x - $width);

        $metrics = $paint->FontMetrics;
        if ($vAlign === VAlign::Middle) {
            $baseline = $y - ($metrics->Ascent + $metrics->Descent) / 2.0;
        } elseif ($vAlign === VAlign::Top) {
            $baseline = $y - $metrics->Ascent;
        } else {
            $baseline = $y;
        }

        if ($outlineColor !== null) {
            $this->TryRecordRuns($page, $clipRect, $fontContext, $text, $drawX, $baseline, $paint->Color, true);
            $canvas->Save();
            $canvas->ClipRect($clipRect);
            self::DrawRuns($canvas, $fontContext, $text, $drawX, $baseline, $paint);
            $canvas->Restore();
        } elseif (!$this->TryRecordRuns($page, $clipRect, $fontContext, $text, $drawX, $baseline, $paint->Color)) {
            self::DrawRuns($canvas, $fontContext, $text, $drawX, $baseline, $paint);
        }
        $this->DrawTextOutline($canvas, $clipRect, $fontContext, $text, $drawX, $baseline, $fontValue, $outlineColor);

        $paint->Dispose();
    }

    private function DrawTitleText($canvas, $page, $clipRect, $text, $x, $y, $fontValue, $color, $hAlign, $vAlign, $outlineColor)
    {
        if ($canvas === null || $text === null || $text === '' || $fontValue === null) {
            return;
        }

        $fontContext = $this->context->GetFontContext($fontValue);
        $paint = $this->TextPaint($fontValue, $color);

        $width = $fontContext->Measure($text);
        $drawX = $hAlign === TextAlign::Left ? $x : ($hAlign === TextAlign::Center ? $x - $width / 2.0 : $x - $width);

        $metrics = $paint->FontMetrics;
        if ($vAlign === VAlign::Middle) {
            $baseline = $y - ($metrics->Ascent + $metrics->Descent) / 2.0;
        } elseif ($vAlign === VAlign::Top) {
            $baseline = $y - $metrics->Ascent;
        } else {
            $baseline = $y;
        }

        $canvas->Save();
        $canvas->ClipRect($clipRect);
        if ($outlineColor !== null) {
            $this->TryRecordRuns($page, $clipRect, $fontContext, $text, $drawX, $baseline, $paint->Color, true);
            self::DrawRuns($canvas, $fontContext, $text, $drawX, $baseline, $paint);
        } elseif (!$this->TryRecordRuns($page, $clipRect, $fontContext, $text, $drawX, $baseline, $paint->Color)) {
            self::DrawRuns($canvas, $fontContext, $text, $drawX, $baseline, $paint);
        }
        $this->DrawTextOutline($canvas, $clipRect, $fontContext, $text, $drawX, $baseline, $fontValue, $outlineColor);
        $canvas->Restore();

        $paint->Dispose();
    }

    private function DrawTextOutline($canvas, $clipRect, $fontContext, $text, $x, $baseline, $fontValue, $outlineColor)
    {
        if ($outlineColor === null) {
            return;
        }

        $paint = $this->TextPaint($fontValue, $outlineColor);
        $paint->Style = SKPaintStyle::Stroke;
        $paint->StrokeWidth = 0.75;
        $canvas->Save();
        $canvas->ClipRect($clipRect);
        self::DrawRuns($canvas, $fontContext, $text, $x, $baseline, $paint);
        $canvas->Restore();
        $paint->Dispose();
    }

    private function DrawRotatedText($canvas, $page, $clipRect, $text, $x, $y, $sizePt, $bold, $color, $hAlign, $vAlign, $angleDegrees)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = $this->ChartFont($sizePt, $bold);
        $fontContext = $this->context->GetFontContext($fontValue);
        $paint = $this->TextPaint($fontValue, $color);

        $width = $fontContext->Measure($text);
        $drawX = $hAlign === TextAlign::Left ? 0.0 : ($hAlign === TextAlign::Center ? -$width / 2.0 : -$width);

        $metrics = $paint->FontMetrics;
        if ($vAlign === VAlign::Middle) {
            $baseline = -($metrics->Ascent + $metrics->Descent) / 2.0;
        } elseif ($vAlign === VAlign::Top) {
            $baseline = -$metrics->Ascent;
        } else {
            $baseline = 0.0;
        }

        if (!$this->TryRecordRotatedRuns($page, $clipRect, $fontContext, $text, $x, $y, $drawX, $baseline, $angleDegrees, $paint->Color)) {
            $canvas->Save();
            $canvas->ClipRect($clipRect);
            $canvas->Translate($x, $y);
            $canvas->RotateDegrees($angleDegrees);
            self::DrawRuns($canvas, $fontContext, $text, $drawX, $baseline, $paint);
            $canvas->Restore();
        }

        $paint->Dispose();
    }

    private function TryRecordRuns($page, $clipRect, $fontContext, $text, $x, $baseline, $color, $invisible = false)
    {
        $session = $this->context->PdfTextSession;
        if ($session === null || $page === null || $text === null || $text === '') {
            return false;
        }

        // The clip can cover a whole chart even for a single-line date label.
        // Its size must not decide whether the text can share font glyphs.
        if (!self::IsTokenizableChartText($text)) {
            return false;
        }

        $runs = $fontContext->SplitRuns($text);
        $cursor = $x;
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $session->RecordChartRun($page, $clipRect, $run->Text, $run->Typeface, $color,
                $fontContext->SizePt, $cursor, $baseline, $invisible);
            $cursor += $run->WidthPt;
        }

        return true;
    }

    private function TryRecordRotatedRuns($page, $clipRect, $fontContext, $text, $originX, $originY, $x, $baseline, $rotationDeg, $color)
    {
        $session = $this->context->PdfTextSession;
        if ($session === null || $page === null || $text === null || $text === '') {
            return false;
        }

        if (!self::IsTokenizableChartText($text)) {
            return false;
        }

        $runs = $fontContext->SplitRuns($text);
        $cursor = $x;
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $session->RecordChartRotatedRun($page, $clipRect, $run->Text, $run->Typeface, $color, $fontContext->SizePt, $originX, $originY, $cursor, $baseline, $rotationDeg);
            $cursor += $run->WidthPt;
        }

        return true;
    }

    private static function IsTokenizableChartText($text)
    {
        if ($text === null || $text === '') {
            return false;
        }

        for ($i = 0; $i < strlen($text); $i++) {
            $ch = $text[$i];
            if ($ch === "\r" || $ch === "\n") {
                return false;
            }
        }

        return true;
    }

    private static function DrawRuns($canvas, $fontContext, $text, $x, $baseline, $paint)
    {
        $runs = $fontContext->SplitRuns($text);
        $cursor = $x;

        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $font = new SKFont($run->Typeface, $fontContext->SizePt);
            $font->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $run->Text, $cursor, $baseline, $font,
                $paint->Color, $paint->Style, $paint->StrokeWidth);
            $font->Dispose();

            $cursor += $run->WidthPt;
        }
    }

    private function TextPaint($font, $color)
    {
        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->IsAntialias = true;
        $paint->Typeface = $this->context->Fonts->Resolve($font);
        $paint->TextSize = (float)$font->Size;
        $paint->TextAlign = SKTextAlign::Left;
        $paint->SubpixelText = true;
        return $paint;
    }

    private function ResolveTitleFont($chart)
    {
        if ($chart !== null && $chart->TitleFont !== null) {
            return $chart->TitleFont;
        }

        return $this->ChartFont(14.0, true, false);
    }

    private function ChartFont($sizePt, $bold = false, $italic = false)
    {
        $font = new FontValue();
        $font->Name = $this->context->Colors->MinorFontName;
        $font->Size = $sizePt;
        $font->Bold = $bold;
        $font->Italic = $italic;
        return $font;
    }
}
