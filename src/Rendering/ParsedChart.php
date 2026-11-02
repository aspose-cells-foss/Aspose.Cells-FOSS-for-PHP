<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

// A chart's definition reduced to what the renderer needs, parsed from the cached values in the
// chart XML (no formula engine required).
final class ParsedChart
{
    public $Kind = ChartKind::Unsupported;

    public $Culture;

    public $TextSizePt;

    public $Title; // null = no title band

    public $TitleFont;

    public $TitleColor;

    public $TitleTextOutlineColor;

    public $TitleAllCaps = false;

    public $HasTitleBorder = false;

    public $TitleBorderColor;

    public $TitleBorderWidthPt = 0.75;

    public $AxisLabelColor;

    public $AxisTextOutlineColor;

    public $LegendTextColor;

    public $LegendTextOutlineColor;

    public $SeriesAxisLabelColor;

    public $CategoryAxisTitle;

    public $ValueAxisTitle;

    public $HasChartAreaFill = false;

    public $ChartAreaFillColor;

    public $ChartAreaPatternPreset;

    public $ChartAreaPatternForegroundColor;

    public $ChartAreaPatternBackgroundColor;

    public $HasChartAreaBorder = false;

    public $ChartAreaBorderColor;

    public $ChartAreaBorderWidthPt = 1.0;

    public $HasChartAreaGlow = false;

    public $ChartAreaGlowColor;

    public $ChartAreaGlowRadiusPt = 0.0;

    public $ChartAreaSoftEdgeRadiusPt = 0.0;

    public $HasChartAreaGradient = false;

    public $ChartAreaGradientColors = array();

    public $ChartAreaGradientPositions = array();

    public $ChartAreaGradientAngleDeg = 0.0;

    public $LegendPosition; // "b","t","l","r","tr", or null

    public $HasManualLegend = false;

    public $LegendX = 0.0;

    public $LegendY = 0.0;

    public $LegendW = 0.0;

    public $LegendH = 0.0;

    public $HasValueGridlines = false;

    public $ValueGridlineColor;

    public $ValueGridlineWidthPt = 1.0;

    public $HasMinorValueGridlines = false;

    public $HasValueAxisLine = false;

    public $ValueAxisOnRight = false;

    public $ValueAxisLineColor;

    public $ValueAxisLineWidthPt = 0.75;

    public $ValueAxisMajorTickMark;

    public $ValueFormatCode;

    public $CrossBetween = true; // category points sit at band centers (Excel default)

    public $IsStacked = false;

    public $IsPercentStacked = false;

    public $AxisMin; // explicit scaling, else auto

    public $AxisMax;

    public $MajorUnit;

    public $MinorUnit;

    public $ScatterXMin;

    public $ScatterXMax;

    public $ScatterXMajorUnit;

    public $ScatterXFormatCode;

    public $ScatterVaryColors = false;

    public $IsBubble = false;

    public $BubbleScalePercent = 100.0;

    public $ShowNegativeBubbles = false;

    public $ScatterValueAxisDeleted = false;

    public $ScatterCategoryAxisDeleted = false;

    public $ScatterTitleOverlay = false;

    public $RoundedCorners = false;

    public $MinorValueGridlineColor;

    public $MinorValueGridlineWidthPt = 1.0;

    public $HasCategoryGridlines = false;

    public $CategoryGridlineColor;

    public $CategoryGridlineWidthPt = 1.0;

    public $HasPlotAreaBorder = false;

    public $PlotAreaBorderColor;

    public $PlotAreaBorderWidthPt = 1.0;

    public $HasDataTable = false;

    public $DataTableShowHorizontalBorders = false;

    public $DataTableShowVerticalBorders = false;

    public $DataTableShowOutline = false;

    public $DataTableShowKeys = false;

    public $DataTableLineColor;

    public $DataTableLineWidthPt = 0.75;

    public $DataTableTextColor;

    public $Series = array();

    public $Categories = array();

    public $CategoryValues = array();

    public $PiePointColors = array();

    public $PiePointLineColors = array();

    public $PiePointLineWidthsPt = array();

    public $IsDoughnut = false;

    public $DoughnutHoleSizePercent = 50.0;

    public $DoughnutFirstSliceAngleDeg = 0.0;

    public $IsOfPie = false;

    public $OfPieType;

    public $OfPieSplitType;

    public $OfPieSplitPosition;

    public $OfPieSecondPieSizePercent = 75.0;

    public $OfPieGapWidthPercent = 150.0;

    public $HasOfPieSeparatorLines = false;

    public $OfPieSeparatorLineColor;

    public $OfPieSeparatorLineWidthPt = 0.75;

    public $HasDateCategoryAxis = false;

    public $CategoryFormatCode;

    public $CategoryBaseTimeUnit;

    public $CategoryAxisTextRotationDeg = 0.0;

    public $CategoryAxisMajorTickMark;

    public $HasCategoryAxisLine = true;

    public $CategoryAxisLineColor;

    public $CategoryAxisLineWidthPt = 1.0;

    public $CategoryAxisAutomaticTextRotation = false;

    public $CategoryAxisMinMax = true;

    public $RadarStyle = 'standard';

    public $GapWidthPercent = 150.0;

    public $OverlapPercent = 0.0;

    public $BarShape = 'box';

    public $IsLine3D = false;

    public $IsArea3D = false;

    public $IsBar3D = false;

    public $IsPie3D = false;

    public $ValueDisplayUnitDivisor = 1.0;

    public $ValueDisplayUnitLabel;

    public $HasManualValueDisplayUnitLabel = false;

    public $ValueDisplayUnitLabelX = 0.0;

    public $ValueDisplayUnitLabelY = 0.0;

    public $RotationXDeg = 15.0;

    public $RotationYDeg = 20.0;

    public $PerspectiveDeg = 30.0;

    public $DepthPercent = 100.0;

    public $GapDepthPercent = 150.0;

    public $RightAngleAxes = true;

    public $HasVisibleSeriesAxis = false;

    public $SideWallImageData;

    public $BackWallImageData;

    public $HasManualTitle = false;

    public $TitleX = 0.0;

    public $TitleY = 0.0;

    // Manual inner-plot layout as fractions of the chart area (Excel c:manualLayout), if present.
    public $HasManualPlot = false;

    public $PlotX = 0.0;

    public $PlotY = 0.0;

    public $PlotW = 0.0;

    public $PlotH = 0.0;

    public function __construct()
    {
        $this->TitleFont = new FontValue();
        $this->TitleFont->Name = 'Calibri';
        $this->TitleFont->Size = 14.0;
        $this->TitleFont->Bold = true;
        $this->TitleColor = new SKColor(0x59, 0x59, 0x59);
        $this->TitleBorderColor = SKColors::Transparent();
        $this->AxisLabelColor = new SKColor(0x59, 0x59, 0x59);
        $this->LegendTextColor = new SKColor(0x59, 0x59, 0x59);
        $this->SeriesAxisLabelColor = new SKColor(0x59, 0x59, 0x59);
        $this->ChartAreaFillColor = SKColors::White();
        $this->ChartAreaPatternForegroundColor = SKColors::Transparent();
        $this->ChartAreaPatternBackgroundColor = SKColors::White();
        $this->ChartAreaBorderColor = new SKColor(0xD9, 0xD9, 0xD9);
        $this->ChartAreaGlowColor = SKColors::Transparent();
        $this->ValueGridlineColor = new SKColor(0xD9, 0xD9, 0xD9);
        $this->ValueAxisLineColor = new SKColor(0x59, 0x59, 0x59);
        $this->MinorValueGridlineColor = new SKColor(0xED, 0xED, 0xED);
        $this->CategoryGridlineColor = new SKColor(0xD9, 0xD9, 0xD9);
        $this->PlotAreaBorderColor = SKColors::Transparent();
        $this->DataTableLineColor = new SKColor(0xD9, 0xD9, 0xD9);
        $this->DataTableTextColor = new SKColor(0x59, 0x59, 0x59);
        $this->OfPieSeparatorLineColor = new SKColor(0xA6, 0xA6, 0xA6);
        $this->CategoryAxisLineColor = new SKColor(0xD9, 0xD9, 0xD9);
    }
}
