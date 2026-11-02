<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\NumberFormatValue;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\DisplayTextFormatter;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class ChartXmlParser
{
    const C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}';

    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function Parse($xml, $colors, $dateSystem, $culture, $workbook = null, $chartModel = null)
    {
        if ($xml === null || $xml === '') {
            return null;
        }

        $doc = null;
        try {
            $doc = XDocument::Parse($xml);
        } catch (\Throwable $exception) {
            return null;
        }

        $chart = $doc->Root !== null ? $doc->Root->Element(self::C . 'chart') : null;
        $plotArea = $chart !== null ? $chart->Element(self::C . 'plotArea') : null;
        if ($plotArea === null) {
            return null;
        }

        $kind = ChartKind::Unsupported;
        $typeEl = self::FindTypeElement($plotArea, $kind);
        if ($typeEl === null) {
            return null;
        }

        $result = new ParsedChart();
        $result->Kind = $kind;
        $result->IsBubble = $typeEl->Name === self::C . 'bubbleChart';
        if ($result->IsBubble) {
            $bubbleScale = 0.0;
            if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'bubbleScale'), 'val'), $bubbleScale)
                && $bubbleScale > 0.0) {
                $result->BubbleScalePercent = $bubbleScale;
            }
            $result->ShowNegativeBubbles = self::AttrOf($typeEl->Element(self::C . 'showNegBubbles'), 'val') === '1';
        }
        $result->RoundedCorners = self::AttrOf($doc->Root->Element(self::C . 'roundedCorners'), 'val') === '1';
        if ($kind === ChartKind::Scatter) {
            $result->ScatterVaryColors = self::AttrOf($typeEl->Element(self::C . 'varyColors'), 'val') === '1';
        }
        $scatterStyle = self::AttrOf($typeEl->Element(self::C . 'scatterStyle'), 'val');
        $smoothScatterStyle = $scatterStyle === 'smooth'
            || $scatterStyle === 'smoothMarker';
        $result->Culture = self::ResolveChartCulture($doc->Root, $culture);
        ChartExplicitFormatting::ApplyTextDefaults($doc->Root, $colors, $result);
        $result->Title = self::ParseTitle($chart, $colors, $result, $chartModel, $workbook);
        self::ParseChartAreaStyle($doc->Root, $colors, $result);
        self::ParseLegend($chart, $colors, $result, $chartModel);
        self::ParseManualPlot($plotArea, $result);
        self::ParseThreeDimensionalSettings($chart, $typeEl, $result, $chartModel);
        self::ParseDoughnutSettings($typeEl, $result);
        self::ParseOfPieSettings($typeEl, $colors, $result);
        self::ParseGrouping($typeEl, $result);
        self::ParseRadarSettings($typeEl, $result);
        self::ParseBarSeriesLayout($typeEl, $result);
        self::ParseValueAxis($plotArea, $colors, $result, $chartModel, $workbook);
        self::ParseCategoryAxis($plotArea, $colors, $result, $chartModel, $workbook);
        self::ParseSeriesAxis($plotArea, $colors, $result, $chartModel);
        self::ParsePlotAreaStyle($plotArea, $colors, $result);
        self::ParseDataTable($plotArea, $colors, $result, $chartModel);
        ChartLegacyGridStyle::Apply($doc->Root, $plotArea, $chartModel, $result);

        $chartDataLabels = $typeEl->Element(self::C . 'dLbls');
        $seriesIndex = 0;
        foreach ($typeEl->Elements(self::C . 'ser') as $ser) {
            $seriesDataLabels = $ser->Element(self::C . 'dLbls') ?? $chartDataLabels;
            $valueHolder = $ser->Element(self::C . 'val') ?? $ser->Element(self::C . 'yVal');
            $series = new ChartSeries();
            $series->Name = self::ParseSeriesName($ser->Element(self::C . 'tx'), $workbook);
            $series->Color = self::ParseSeriesColor($ser, $colors, $seriesIndex);
            $series->LineWidthPt = self::ParseLineWidthPt($ser);
            $smoothAttribute = self::AttrOf($ser->Element(self::C . 'smooth'), 'val');
            $series->SmoothLine = $kind === ChartKind::Scatter && ($ser->Element(self::C . 'smooth') === null
                ? $smoothScatterStyle
                : $smoothAttribute === '1');
            $series->Values = self::ParseNumericData($valueHolder, $workbook, $dateSystem);
            $series->XValues = self::ParseNumericData($ser->Element(self::C . 'xVal'), $workbook, $dateSystem);
            $series->BubbleSizes = self::ParseNumericData($ser->Element(self::C . 'bubbleSize'), $workbook, $dateSystem);
            $series->Bubble3D = self::AttrOf($ser->Element(self::C . 'bubble3D'), 'val') === '1';
            $series->ShowValueDataLabels = self::AttrOf(
                $seriesDataLabels !== null ? $seriesDataLabels->Element(self::C . 'showVal') : null,
                'val') === '1';
            $series->DataLabelPosition = self::AttrOf($seriesDataLabels !== null ? $seriesDataLabels->Element(self::C . 'dLblPos') : null, 'val');
            $series->ValueFormatCode = self::TextOf(self::FirstDescendant($valueHolder, self::C . 'formatCode'));
            ChartExplicitFormatting::ApplySeries($ser, $typeEl, $colors, $series, $seriesIndex);
            self::ParseSeriesLineStyle($ser, $colors, $series);
            if (($kind === ChartKind::Column || $kind === ChartKind::Bar) && !$result->IsBar3D) {
                ChartBarBorder::Apply($ser, $series);
            }
            self::ParseSeriesPatternFill($ser, $colors, $series);
            self::ParseMarker($ser, $colors, $series);
            self::ParseErrorBars($ser, $colors, $series);
            $dataLabelColor = null;
            if (self::TryResolveTextColor($seriesDataLabels !== null ? $seriesDataLabels->Element(self::C . 'txPr') : null,
                $colors, $chartModel, $dataLabelColor)) {
                $series->DataLabelColor = $dataLabelColor;
            }
            self::ApplyRadarSeriesDefaults($result, $series);
            $result->Series[] = $series;
            $seriesIndex++;
        }

        self::ApplyFallbackTitle($result);

        $firstSer = self::FirstElement($typeEl, self::C . 'ser');
        if ($firstSer !== null) {
            $result->Categories = self::ParseCategories($firstSer->Element(self::C . 'cat'), $workbook, $dateSystem, $result->Culture);
            $result->CategoryValues = self::ParseCategoryValues($firstSer->Element(self::C . 'cat'), $workbook, $dateSystem);
            PieChartPointMetadata::Apply($firstSer, $colors, $result);
            $result->PiePointColors = self::ParsePiePointColors($firstSer, $colors, count($result->Categories));
            if ($result->ValueFormatCode === null || $result->ValueFormatCode === '') {
                $result->ValueFormatCode = self::TextOf(self::FirstDescendant($firstSer->Element(self::C . 'val'), self::C . 'formatCode'));
            }
        }

        return $result;
    }

    private static function ParseDataTable($plotArea, $colors, $result, $chartModel)
    {
        $dataTable = $plotArea !== null ? $plotArea->Element(self::C . 'dTable') : null;
        if ($dataTable === null || $result === null) {
            return;
        }

        $result->HasDataTable = true;
        $result->DataTableShowHorizontalBorders = self::IsEnabled($dataTable->Element(self::C . 'showHorzBorder'));
        $result->DataTableShowVerticalBorders = self::IsEnabled($dataTable->Element(self::C . 'showVertBorder'));
        $result->DataTableShowOutline = self::IsEnabled($dataTable->Element(self::C . 'showOutline'));
        $result->DataTableShowKeys = self::IsEnabled($dataTable->Element(self::C . 'showKeys'));

        $line = $dataTable->Element(self::C . 'spPr') !== null
            ? $dataTable->Element(self::C . 'spPr')->Element(self::A . 'ln')
            : null;
        $hasVisibleLine = false;
        $lineColor = null;
        $lineWidth = 0.0;
        self::ParseLineStyle($line, $colors, $hasVisibleLine, $lineColor, $lineWidth);
        if ($hasVisibleLine) {
            $result->DataTableLineColor = $lineColor;
            $result->DataTableLineWidthPt = $lineWidth;
        }

        $textColor = null;
        if (self::TryResolveTextColor($dataTable->Element(self::C . 'txPr'), $colors, $chartModel, $textColor)) {
            $result->DataTableTextColor = $textColor;
        }
    }

    private static function ParseErrorBars($seriesElement, $colors, $series)
    {
        $errorBars = $seriesElement !== null ? $seriesElement->Element(self::C . 'errBars') : null;
        if ($errorBars === null || $series === null) {
            return;
        }

        $series->HasErrorBars = true;
        $series->ErrorBarType = self::AttrOf($errorBars->Element(self::C . 'errBarType'), 'val');
        $series->ErrorValueType = self::AttrOf($errorBars->Element(self::C . 'errValType'), 'val');
        $series->ErrorBarsHaveEndCaps = self::AttrOf($errorBars->Element(self::C . 'noEndCap'), 'val') !== '1';

        $line = $errorBars->Element(self::C . 'spPr') !== null
            ? $errorBars->Element(self::C . 'spPr')->Element(self::A . 'ln')
            : null;
        $hasVisibleLine = false;
        $lineColor = null;
        $lineWidth = 0.0;
        self::ParseLineStyle($line, $colors, $hasVisibleLine, $lineColor, $lineWidth);
        if ($hasVisibleLine) {
            $series->ErrorBarColor = $lineColor;
            $series->ErrorBarWidthPt = $lineWidth;
        }
    }

    private static function IsEnabled($element)
    {
        return $element !== null
            && self::AttrOf($element, 'val') !== '0';
    }

    private static function ApplyFallbackTitle($chart)
    {
        if ($chart === null) {
            return;
        }

        if (($chart->Kind === ChartKind::Pie && !$chart->IsDoughnut
                || $chart->Kind === ChartKind::Scatter && $chart->Series !== null && count($chart->Series) === 1)
            && $chart->Series !== null
            && count($chart->Series) > 0
            && ($chart->Title === null || $chart->Title === '' || $chart->Title === 'Chart Title')) {
            $chart->Title = $chart->Series[0]->Name;
        }
    }

    private static function ResolveChartCulture($chartSpace, $fallbackCulture)
    {
        $cultureName = self::AttrOf($chartSpace !== null ? $chartSpace->Element(self::C . 'lang') : null, 'val');
        if ($cultureName === null || $cultureName === '') {
            return $fallbackCulture !== null ? $fallbackCulture : CultureInfo::GetInvariantCulture();
        }

        try {
            return CultureInfo::GetCultureInfo($cultureName);
        } catch (\Throwable $exception) {
            return $fallbackCulture !== null ? $fallbackCulture : CultureInfo::GetInvariantCulture();
        }
    }

    private static function FindTypeElement($plotArea, &$kind)
    {
        $line = $plotArea->Element(self::C . 'lineChart');
        if ($line !== null) { $kind = ChartKind::Line; return $line; }

        $scatter = $plotArea->Element(self::C . 'scatterChart');
        if ($scatter !== null) { $kind = ChartKind::Scatter; return $scatter; }

        $bubble = $plotArea->Element(self::C . 'bubbleChart');
        if ($bubble !== null) { $kind = ChartKind::Scatter; return $bubble; }

        $line3d = $plotArea->Element(self::C . 'line3DChart');
        if ($line3d !== null) { $kind = ChartKind::Line; return $line3d; }

        $bar = $plotArea->Element(self::C . 'barChart');
        if ($bar !== null) {
            $dir = self::AttrOf($bar->Element(self::C . 'barDir'), 'val');
            $kind = $dir === 'bar' ? ChartKind::Bar : ChartKind::Column;
            return $bar;
        }

        $bar3d = $plotArea->Element(self::C . 'bar3DChart');
        if ($bar3d !== null) {
            $dir = self::AttrOf($bar3d->Element(self::C . 'barDir'), 'val');
            $kind = $dir === 'bar' ? ChartKind::Bar : ChartKind::Column;
            return $bar3d;
        }

        $area = $plotArea->Element(self::C . 'areaChart');
        if ($area !== null) { $kind = ChartKind::Area; return $area; }

        $area3d = $plotArea->Element(self::C . 'area3DChart');
        if ($area3d !== null) { $kind = ChartKind::Area; return $area3d; }

        $radar = $plotArea->Element(self::C . 'radarChart');
        if ($radar !== null) { $kind = ChartKind::Radar; return $radar; }

        $pie = $plotArea->Element(self::C . 'pieChart');
        if ($pie !== null) { $kind = ChartKind::Pie; return $pie; }

        $pie3d = $plotArea->Element(self::C . 'pie3DChart');
        if ($pie3d !== null) { $kind = ChartKind::Pie; return $pie3d; }

        $doughnut = $plotArea->Element(self::C . 'doughnutChart');
        if ($doughnut !== null) { $kind = ChartKind::Pie; return $doughnut; }

        $ofPie = $plotArea->Element(self::C . 'ofPieChart');
        if ($ofPie !== null) { $kind = ChartKind::Pie; return $ofPie; }

        $kind = ChartKind::Unsupported;
        return null;
    }

    private static function ParseDoughnutSettings($typeEl, $result)
    {
        if ($typeEl === null || $typeEl->Name !== self::C . 'doughnutChart') {
            return;
        }

        $result->IsDoughnut = true;
        $value = 0.0;
        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'holeSize'), 'val'), $value)) {
            $result->DoughnutHoleSizePercent = max(10.0, min(90.0, $value));
        }
        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'firstSliceAng'), 'val'), $value)) {
            $result->DoughnutFirstSliceAngleDeg = $value;
        }
    }

    private static function ParseTitle($chart, $colors, $result, $chartModel, $workbook)
    {
        $titleEl = $chart->Element(self::C . 'title');
        if ($titleEl === null) {
            return null;
        }

        if (self::AttrOf($chart->Element(self::C . 'autoTitleDeleted'), 'val') === '1') {
            return null;
        }

        self::ParseManualTitleLayout($titleEl, $result);
        $result->ScatterTitleOverlay = $result->Kind === ChartKind::Scatter
            && self::AttrOf($titleEl->Element(self::C . 'overlay'), 'val') === '1';
        self::ParseTitleStyle($titleEl, $colors, $result, $chartModel);
        $text = self::ParseTitleText($titleEl, $workbook);
        if ($text !== null && $text !== '') {
            return $text;
        }

        return 'Chart Title';
    }

    private static function ParseManualTitleLayout($titleEl, $result)
    {
        $manual = $titleEl->Element(self::C . 'layout') !== null
            ? $titleEl->Element(self::C . 'layout')->Element(self::C . 'manualLayout')
            : null;
        if ($manual === null) {
            return;
        }

        $x = 0.0;
        $y = 0.0;
        if (self::TryDouble(self::AttrOf($manual->Element(self::C . 'x'), 'val'), $x)
            && self::TryDouble(self::AttrOf($manual->Element(self::C . 'y'), 'val'), $y)) {
            $result->HasManualTitle = true;
            $result->TitleX = $x;
            $result->TitleY = $y;
        }
    }

    private static function ParseTitleStyle($titleEl, $colors, $result, $chartModel)
    {
        if ($titleEl === null || $result === null) {
            return;
        }

        $titleTextProperties = self::ResolveTextProperties($titleEl);

        $defaultRunProperties = $titleTextProperties !== null
            ? self::FirstDescendant($titleTextProperties, self::A . 'defRPr')
            : null;
        $result->TitleAllCaps = self::AttrOf($defaultRunProperties, 'cap') === 'all';

        $titleColor = null;
        if (self::TryResolveTextColor($titleTextProperties, $colors, $chartModel, $titleColor)) {
            $result->TitleColor = $titleColor;
        }
        if (self::TryResolveTextOutlineColor($titleTextProperties, $colors, $titleColor)) {
            $result->TitleTextOutlineColor = $titleColor;
        }

        $titleFont = null;
        if (self::TryResolveTextFont($titleTextProperties, $colors, $titleFont)) {
            $result->TitleFont = $titleFont;
        }

        $shapeProperties = $titleEl->Element(self::C . 'spPr');
        $line = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'ln') : null;
        self::ParseLineStyle($line, $colors, $result->HasTitleBorder, $result->TitleBorderColor, $result->TitleBorderWidthPt);
        if ($result->HasTitleBorder && $line !== null && $line->Attribute('w') === null) {
            $result->TitleBorderWidthPt = 0.75;
        }
    }

    private static function ParseLegend($chart, $colors, $result, $chartModel)
    {
        $legend = $chart->Element(self::C . 'legend');
        if ($legend !== null) {
            $result->LegendPosition = self::AttrOf($legend->Element(self::C . 'legendPos'), 'val') ?? 'r';
            $layout = $legend->Element(self::C . 'layout');
            $manual = $layout !== null ? $layout->Element(self::C . 'manualLayout') : null;
            $x = 0.0;
            $y = 0.0;
            $w = 0.0;
            $h = 0.0;
            if ($manual !== null
                && self::TryDouble(self::AttrOf($manual->Element(self::C . 'x'), 'val'), $x)
                && self::TryDouble(self::AttrOf($manual->Element(self::C . 'y'), 'val'), $y)
                && self::TryDouble(self::AttrOf($manual->Element(self::C . 'w'), 'val'), $w)
                && self::TryDouble(self::AttrOf($manual->Element(self::C . 'h'), 'val'), $h)) {
                $result->HasManualLegend = true;
                $result->LegendX = $x;
                $result->LegendY = $y;
                $result->LegendW = $w;
                $result->LegendH = $h;
            }

            $legendTextColor = null;
            if (self::TryResolveTextColor($legend->Element(self::C . 'txPr'), $colors, $chartModel, $legendTextColor)) {
                $result->LegendTextColor = $legendTextColor;
            }
            if (self::TryResolveTextOutlineColor($legend->Element(self::C . 'txPr'), $colors, $legendTextColor)) {
                $result->LegendTextOutlineColor = $legendTextColor;
            }
        }
    }

    private static function ParseGrouping($typeEl, $result)
    {
        if ($typeEl === null || $result === null) {
            return;
        }

        $grouping = self::AttrOf($typeEl->Element(self::C . 'grouping'), 'val');
        $result->IsStacked = $grouping === 'stacked';
        $result->IsPercentStacked = $grouping === 'percentStacked';
    }

    private static function ParseBarSeriesLayout($typeEl, $result)
    {
        if ($typeEl === null || $result === null) {
            return;
        }

        if ($typeEl->Name !== self::C . 'barChart' && $typeEl->Name !== self::C . 'bar3DChart') {
            return;
        }

        $gapWidth = 0.0;
        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'gapWidth'), 'val'), $gapWidth) && $gapWidth >= 0.0) {
            $result->GapWidthPercent = $gapWidth;
        }

        $overlap = 0.0;
        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'overlap'), 'val'), $overlap)) {
            if ($overlap < -100.0) {
                $overlap = -100.0;
            } elseif ($overlap > 100.0) {
                $overlap = 100.0;
            }

            $result->OverlapPercent = $overlap;
        }

        $shape = self::AttrOf($typeEl->Element(self::C . 'shape'), 'val');
        if ($shape !== null && $shape !== '') {
            $result->BarShape = $shape;
        }
    }

    private static function ParseRadarSettings($typeEl, $result)
    {
        if ($typeEl === null || $result === null) {
            return;
        }

        if ($typeEl->Name !== self::C . 'radarChart') {
            return;
        }

        $radarStyle = self::AttrOf($typeEl->Element(self::C . 'radarStyle'), 'val');
        if ($radarStyle !== null && $radarStyle !== '') {
            $result->RadarStyle = $radarStyle;
        }
    }

    private static function ParseOfPieSettings($typeEl, $colors, $result)
    {
        if ($typeEl === null || $result === null) {
            return;
        }

        if ($typeEl->Name !== self::C . 'ofPieChart') {
            return;
        }

        $result->IsOfPie = true;
        $result->OfPieType = self::AttrOf($typeEl->Element(self::C . 'ofPieType'), 'val') ?? 'pie';
        $result->OfPieSplitType = self::AttrOf($typeEl->Element(self::C . 'splitType'), 'val');

        $value = 0.0;
        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'splitPos'), 'val'), $value)) {
            $result->OfPieSplitPosition = $value;
        }

        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'secondPieSize'), 'val'), $value) && $value > 0.0) {
            $result->OfPieSecondPieSizePercent = $value;
        }

        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'gapWidth'), 'val'), $value) && $value >= 0.0) {
            $result->OfPieGapWidthPercent = $value;
        }

        $separatorLine = $typeEl->Element(self::C . 'serLines');
        if ($separatorLine === null) {
            return;
        }

        $line = $separatorLine->Element(self::C . 'spPr') !== null ? $separatorLine->Element(self::C . 'spPr')->Element(self::A . 'ln') : null;
        $color = null;
        $widthPt = 0.0;
        $hasVisibleLine = false;
        self::ParseLineStyle($line, $colors, $hasVisibleLine, $color, $widthPt);
        $result->HasOfPieSeparatorLines = $hasVisibleLine || $line === null;
        if ($hasVisibleLine) {
            $result->OfPieSeparatorLineColor = $color;
            $result->OfPieSeparatorLineWidthPt = $widthPt;
        }
    }

    private static function ParseThreeDimensionalSettings($chart, $typeEl, $result, $chartModel)
    {
        if ($chart === null || $typeEl === null || $result === null) {
            return;
        }

        $isLine3D = $typeEl->Name === self::C . 'line3DChart';
        $isArea3D = $typeEl->Name === self::C . 'area3DChart';
        $isBar3D = $typeEl->Name === self::C . 'bar3DChart';
        $isPie3D = $typeEl->Name === self::C . 'pie3DChart';
        if (!$isLine3D && !$isArea3D && !$isBar3D && !$isPie3D) {
            return;
        }

        $result->IsLine3D = $isLine3D;
        $result->IsArea3D = $isArea3D;
        $result->IsBar3D = $isBar3D;
        $result->IsPie3D = $isPie3D;
        $result->SideWallImageData = self::ResolveWallImageData($chart->Element(self::C . 'sideWall'), $chartModel);
        $result->BackWallImageData = self::ResolveWallImageData($chart->Element(self::C . 'backWall'), $chartModel);
        $view3D = $chart->Element(self::C . 'view3D');
        $value = 0.0;
        if (self::TryDouble(self::AttrOf($view3D !== null ? $view3D->Element(self::C . 'rotX') : null, 'val'), $value)) {
            $result->RotationXDeg = $value;
        }

        if (self::TryDouble(self::AttrOf($view3D !== null ? $view3D->Element(self::C . 'rotY') : null, 'val'), $value)) {
            $result->RotationYDeg = $value;
        }

        if (self::TryDouble(self::AttrOf($view3D !== null ? $view3D->Element(self::C . 'perspective') : null, 'val'), $value)) {
            $result->PerspectiveDeg = $value;
        }

        if (self::TryDouble(self::AttrOf($view3D !== null ? $view3D->Element(self::C . 'depthPercent') : null, 'val'), $value) && $value > 0.0) {
            $result->DepthPercent = $value;
        }

        $rightAngleAxes = self::AttrOf($view3D !== null ? $view3D->Element(self::C . 'rAngAx') : null, 'val');
        if ($rightAngleAxes !== null && $rightAngleAxes !== '') {
            $result->RightAngleAxes = $rightAngleAxes !== '0'
                && strcasecmp($rightAngleAxes, 'false') !== 0;
        }

        if (self::TryDouble(self::AttrOf($typeEl->Element(self::C . 'gapDepth'), 'val'), $value) && $value >= 0.0) {
            $result->GapDepthPercent = $value;
        }
    }

    private static function ResolveWallImageData($wall, $chartModel)
    {
        if ($wall === null || $chartModel === null) {
            return null;
        }

        $shapeProperties = $wall->Element(self::C . 'spPr');
        $blipFill = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'blipFill') : null;
        $blip = $blipFill !== null ? $blipFill->Element(self::A . 'blip') : null;
        $relationshipId = $blip !== null
            ? (string) $blip->Attribute('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            : null;
        if (($relationshipId === null || $relationshipId === '') || $blipFill->Element(self::A . 'tile') === null) {
            return null;
        }

        for ($i = 0; $i < count($chartModel->CompanionFiles); $i++) {
            $companion = $chartModel->CompanionFiles[$i];
            if ($companion->RelationshipId === $relationshipId
                && $companion->RawContent !== null
                && strlen($companion->RawContent) > 0) {
                return $companion->RawContent;
            }
        }

        return null;
    }

    private static function ParseManualPlot($plotArea, $result)
    {
        $manual = $plotArea->Element(self::C . 'layout') !== null
            ? $plotArea->Element(self::C . 'layout')->Element(self::C . 'manualLayout')
            : null;
        if ($manual === null || self::AttrOf($manual->Element(self::C . 'layoutTarget'), 'val') !== 'inner') {
            return;
        }

        $x = 0.0;
        $y = 0.0;
        $w = 0.0;
        $h = 0.0;
        if (self::TryDouble(self::AttrOf($manual->Element(self::C . 'x'), 'val'), $x)
            && self::TryDouble(self::AttrOf($manual->Element(self::C . 'y'), 'val'), $y)
            && self::TryDouble(self::AttrOf($manual->Element(self::C . 'w'), 'val'), $w)
            && self::TryDouble(self::AttrOf($manual->Element(self::C . 'h'), 'val'), $h)) {
            $result->HasManualPlot = true;
            $result->PlotX = $x;
            $result->PlotY = $y;
            $result->PlotW = $w;
            $result->PlotH = $h;
        }
    }

    private static function ParseValueAxis($plotArea, $colors, $result, $chartModel, $workbook)
    {
        $valAx = $result->Kind === ChartKind::Scatter
            ? self::FindValueAxisAt($plotArea, 'l', 'r')
            : self::FirstElement($plotArea, self::C . 'valAx');
        if ($valAx === null) {
            return;
        }

        $result->HasValueGridlines = $valAx->Element(self::C . 'majorGridlines') !== null;
        $result->HasMinorValueGridlines = $valAx->Element(self::C . 'minorGridlines') !== null;
        $result->ValueAxisMajorTickMark = self::AttrOf($valAx->Element(self::C . 'majorTickMark'), 'val');
        $result->ValueAxisOnRight = self::AttrOf($valAx->Element(self::C . 'axPos'), 'val') === 'r';
        $result->ValueFormatCode = self::AttrOf($valAx->Element(self::C . 'numFmt'), 'formatCode');
        $result->ValueAxisTitle = self::ParseAxisTitle($valAx, $workbook);
        $result->ScatterValueAxisDeleted = $result->Kind === ChartKind::Scatter
            && self::AttrOf($valAx->Element(self::C . 'delete'), 'val') === '1';
        $displayUnits = $valAx->Element(self::C . 'dispUnits');
        $displayUnit = self::AttrOf($displayUnits !== null
            ? $displayUnits->Element(self::C . 'builtInUnit')
            : null, 'val');
        if ($displayUnit === 'hundreds') {
            $result->ValueDisplayUnitDivisor = 100.0;
            $result->ValueDisplayUnitLabel = 'Hundreds';
        }
        $displayUnitLabel = $displayUnits !== null ? $displayUnits->Element(self::C . 'dispUnitsLbl') : null;
        $displayUnitLayout = $displayUnitLabel !== null ? $displayUnitLabel->Element(self::C . 'layout') : null;
        $displayUnitManual = $displayUnitLayout !== null ? $displayUnitLayout->Element(self::C . 'manualLayout') : null;
        $displayUnitX = 0.0;
        $displayUnitY = 0.0;
        if ($displayUnitManual !== null
            && self::TryDouble(self::AttrOf($displayUnitManual->Element(self::C . 'x'), 'val'), $displayUnitX)
            && self::TryDouble(self::AttrOf($displayUnitManual->Element(self::C . 'y'), 'val'), $displayUnitY)) {
            $result->HasManualValueDisplayUnitLabel = true;
            $result->ValueDisplayUnitLabelX = $displayUnitX;
            $result->ValueDisplayUnitLabelY = $displayUnitY;
        }
        $result->CrossBetween = self::AttrOf($valAx->Element(self::C . 'crossBetween'), 'val') !== 'midCat';

        $shapeProperties = $valAx->Element(self::C . 'spPr');
        $axisLine = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'ln') : null;
        self::ParseLineStyle($axisLine, $colors, $result->HasValueAxisLine, $result->ValueAxisLineColor, $result->ValueAxisLineWidthPt);
        if ($axisLine !== null && $axisLine->Attribute('w') === null) {
            $result->ValueAxisLineWidthPt = 0.75;
        }
        self::ParseValueGridlineStyle($valAx->Element(self::C . 'majorGridlines'), $colors, $result, false);
        self::ParseValueGridlineStyle($valAx->Element(self::C . 'minorGridlines'), $colors, $result, true);
        $scaling = $valAx->Element(self::C . 'scaling');
        $v = 0.0;
        if ($scaling !== null) {
            if (self::TryDouble(self::AttrOf($scaling->Element(self::C . 'min'), 'val'), $v)) $result->AxisMin = $v;
            if (self::TryDouble(self::AttrOf($scaling->Element(self::C . 'max'), 'val'), $v)) $result->AxisMax = $v;
        }
        if (self::TryDouble(self::AttrOf($valAx->Element(self::C . 'majorUnit'), 'val'), $v)) $result->MajorUnit = $v;
        if (self::TryDouble(self::AttrOf($valAx->Element(self::C . 'minorUnit'), 'val'), $v)) $result->MinorUnit = $v;

        $axisLabelColor = null;
        if (self::TryResolveTextColor($valAx->Element(self::C . 'txPr'), $colors, $chartModel, $axisLabelColor)) {
            $result->AxisLabelColor = $axisLabelColor;
        }
        if (self::TryResolveTextOutlineColor($valAx->Element(self::C . 'txPr'), $colors, $axisLabelColor)) {
            $result->AxisTextOutlineColor = $axisLabelColor;
        }
    }

    private static function ParseCategoryAxis($plotArea, $colors, $result, $chartModel, $workbook)
    {
        $dateAxis = self::FirstElement($plotArea, self::C . 'dateAx');
        $categoryAxis = self::FirstElement($plotArea, self::C . 'catAx');
        $scatterAxis = $result->Kind === ChartKind::Scatter
            ? self::FindValueAxisAt($plotArea, 'b', 't')
            : null;
        $axisElement = $dateAxis ?? $categoryAxis ?? $scatterAxis;
        if ($axisElement === null) {
            return;
        }

        $result->HasDateCategoryAxis = $dateAxis !== null;
        $result->CategoryFormatCode = self::AttrOf($axisElement->Element(self::C . 'numFmt'), 'formatCode');
        if ($scatterAxis !== null) {
            $result->ScatterXFormatCode = $result->CategoryFormatCode;
            $scatterScaling = $scatterAxis->Element(self::C . 'scaling');
            $scatterValue = 0.0;
            if ($scatterScaling !== null) {
                if (self::TryDouble(self::AttrOf($scatterScaling->Element(self::C . 'min'), 'val'), $scatterValue)) $result->ScatterXMin = $scatterValue;
                if (self::TryDouble(self::AttrOf($scatterScaling->Element(self::C . 'max'), 'val'), $scatterValue)) $result->ScatterXMax = $scatterValue;
            }
            if (self::TryDouble(self::AttrOf($scatterAxis->Element(self::C . 'majorUnit'), 'val'), $scatterValue)) $result->ScatterXMajorUnit = $scatterValue;
        }
        $result->CategoryAxisTitle = self::ParseAxisTitle($axisElement, $workbook);
        $result->ScatterCategoryAxisDeleted = $result->Kind === ChartKind::Scatter
            && self::AttrOf($axisElement->Element(self::C . 'delete'), 'val') === '1';
        $result->CategoryBaseTimeUnit = self::AttrOf($axisElement->Element(self::C . 'baseTimeUnit'), 'val');
        $result->CategoryAxisMajorTickMark = self::AttrOf($axisElement->Element(self::C . 'majorTickMark'), 'val');
        $axisShape = $axisElement->Element(self::C . 'spPr');
        $axisLine = $axisShape !== null ? $axisShape->Element(self::A . 'ln') : null;
        if ($axisLine !== null) {
            self::ParseLineStyle($axisLine, $colors, $result->HasCategoryAxisLine,
                $result->CategoryAxisLineColor, $result->CategoryAxisLineWidthPt);
        }
        $scaling = $axisElement->Element(self::C . 'scaling');
        $result->CategoryAxisMinMax = self::AttrOf($scaling !== null ? $scaling->Element(self::C . 'orientation') : null, 'val') !== 'maxMin';
        self::ParseGridlineStyle($axisElement->Element(self::C . 'majorGridlines'), $colors, $result);

        $bodyPr = $axisElement->Element(self::C . 'txPr') !== null ? $axisElement->Element(self::C . 'txPr')->Element(self::A . 'bodyPr') : null;
        $rotation = 0.0;
        if (self::TryDouble(self::AttrOf($bodyPr, 'rot'), $rotation)) {
            // Excel's automatic orientation marker is not a literal -1000 degree angle.
            $result->CategoryAxisAutomaticTextRotation = $rotation === -60000000.0;
            $result->CategoryAxisTextRotationDeg = max(-90.0, min(90.0, $rotation / 60000.0));
        }

        $axisLabelColor = null;
        if (self::TryResolveTextColor($axisElement->Element(self::C . 'txPr'), $colors, $chartModel, $axisLabelColor)) {
            $result->AxisLabelColor = $axisLabelColor;
        }
        if (self::TryResolveTextOutlineColor($axisElement->Element(self::C . 'txPr'), $colors, $axisLabelColor)) {
            $result->AxisTextOutlineColor = $axisLabelColor;
        }
    }

    private static function FindValueAxisAt($plotArea, $firstPosition, $secondPosition)
    {
        if ($plotArea === null) {
            return null;
        }

        foreach ($plotArea->Elements(self::C . 'valAx') as $axis) {
            $position = self::AttrOf($axis->Element(self::C . 'axPos'), 'val');
            if ($position === $firstPosition
                || $position === $secondPosition) {
                return $axis;
            }
        }

        return null;
    }

    private static function ParsePlotAreaStyle($plotArea, $colors, $result)
    {
        if ($plotArea === null || $colors === null || $result === null) {
            return;
        }

        $shapeProperties = $plotArea->Element(self::C . 'spPr');
        $line = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'ln') : null;
        self::ParseLineStyle($line, $colors, $result->HasPlotAreaBorder, $result->PlotAreaBorderColor, $result->PlotAreaBorderWidthPt);
        if ($result->Kind === ChartKind::Scatter && $line !== null && self::AttrOf($line, 'w') === null) {
            $result->PlotAreaBorderWidthPt = 0.75;
        }
    }

    private static function ParseChartAreaStyle($chartSpace, $colors, $result)
    {
        if ($chartSpace === null || $colors === null || $result === null) {
            return;
        }

        $shapeProperties = $chartSpace->Element(self::C . 'spPr');
        $borderLine = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'ln') : null;
        self::ParseLineStyle($borderLine, $colors, $result->HasChartAreaBorder,
            $result->ChartAreaBorderColor, $result->ChartAreaBorderWidthPt);
        $patternFill = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'pattFill') : null;
        if ($patternFill !== null) {
            $result->ChartAreaPatternPreset = self::AttrOf($patternFill, 'prst');
            $patternColor = null;
            if (self::TryDrawingColor(self::FillChild($patternFill->Element(self::A . 'fgClr')), $colors, $patternColor)) {
                $result->ChartAreaPatternForegroundColor = $patternColor;
            }
            if (self::TryDrawingColor(self::FillChild($patternFill->Element(self::A . 'bgClr')), $colors, $patternColor)) {
                $result->ChartAreaPatternBackgroundColor = $patternColor;
            }
        }

        $effectList = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'effectLst') : null;
        $glow = $effectList !== null ? $effectList->Element(self::A . 'glow') : null;
        $glowRadius = 0.0;
        $glowColor = null;
        if (self::TryDouble(self::AttrOf($glow, 'rad'), $glowRadius)
            && $glowRadius > 0.0
            && self::TryDrawingColor(self::FillChild($glow), $colors, $glowColor)) {
            $result->HasChartAreaGlow = true;
            $result->ChartAreaGlowRadiusPt = $glowRadius / 12700.0;
            $result->ChartAreaGlowColor = $glowColor;
        }
        $softEdge = $effectList !== null ? $effectList->Element(self::A . 'softEdge') : null;
        $softEdgeRadius = 0.0;
        if (self::TryDouble(self::AttrOf($softEdge, 'rad'), $softEdgeRadius) && $softEdgeRadius > 0.0) {
            $result->ChartAreaSoftEdgeRadiusPt = $softEdgeRadius / 12700.0;
        }
        $gradientFill = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'gradFill') : null;
        if ($gradientFill !== null) {
            $stops = $gradientFill->Element(self::A . 'gsLst');
            $parsedStops = array();
            if ($stops !== null) {
                foreach ($stops->Elements(self::A . 'gs') as $stop) {
                    $position = 0.0;
                    $stopColor = null;
                    if (self::TryDouble(self::AttrOf($stop, 'pos'), $position)
                        && self::TryDrawingColor(self::FillChild($stop), $colors, $stopColor)) {
                        $parsedStops[] = array((float)($position / 100000.0), $stopColor);
                    }
                }
            }

            usort($parsedStops, function ($left, $right) {
                return $left[0] <=> $right[0];
            });
            for ($i = 0; $i < count($parsedStops); $i++) {
                $result->ChartAreaGradientPositions[] = $parsedStops[$i][0];
                $result->ChartAreaGradientColors[] = $parsedStops[$i][1];
            }

            $angle = 0.0;
            $linear = $gradientFill->Element(self::A . 'lin');
            if (self::TryDouble(self::AttrOf($linear, 'ang'), $angle)) {
                $result->ChartAreaGradientAngleDeg = $angle / 60000.0;
            }

            $result->HasChartAreaGradient = count($result->ChartAreaGradientColors) >= 2;
            if ($result->HasChartAreaGradient) {
                return;
            }
        }

        $solidFill = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'solidFill') : null;
        $fill = self::FillChild($solidFill);
        if ($fill === null) {
            return;
        }

        $color = null;
        if (self::TryDrawingColor($fill, $colors, $color)) {
            $result->HasChartAreaFill = true;
            $result->ChartAreaFillColor = $color;
        }
    }

    private static function ParseSeriesAxis($plotArea, $colors, $result, $chartModel)
    {
        $serAx = self::FirstElement($plotArea, self::C . 'serAx');
        if ($serAx === null) {
            return;
        }

        $result->HasVisibleSeriesAxis = self::AttrOf($serAx->Element(self::C . 'delete'), 'val') !== '1';
        $axisLabelColor = null;
        if (self::TryResolveTextColor($serAx->Element(self::C . 'txPr'), $colors, $chartModel, $axisLabelColor)) {
            $result->SeriesAxisLabelColor = $axisLabelColor;
        }
    }

    private static function ParseGridlineStyle($gridlines, $colors, $result)
    {
        if ($gridlines === null || $result === null) {
            return;
        }

        $line = $gridlines->Element(self::C . 'spPr') !== null ? $gridlines->Element(self::C . 'spPr')->Element(self::A . 'ln') : null;
        $color = null;
        $widthPt = 0.0;
        $hasVisibleLine = false;
        self::ParseLineStyle($line, $colors, $hasVisibleLine, $color, $widthPt);
        $result->HasCategoryGridlines = $hasVisibleLine || $line === null;
        if ($hasVisibleLine) {
            $result->CategoryGridlineColor = $color;
            $result->CategoryGridlineWidthPt = $widthPt;
        }
    }

    private static function ParseValueGridlineStyle($gridlines, $colors, $result, $minor)
    {
        if ($gridlines === null || $result === null) {
            return;
        }

        $line = $gridlines->Element(self::C . 'spPr') !== null ? $gridlines->Element(self::C . 'spPr')->Element(self::A . 'ln') : null;
        $color = null;
        $widthPt = 0.0;
        $hasVisibleLine = false;
        self::ParseLineStyle($line, $colors, $hasVisibleLine, $color, $widthPt);
        if ($minor) {
            $result->HasMinorValueGridlines = $hasVisibleLine || $line === null;
            if ($hasVisibleLine) {
                $result->MinorValueGridlineColor = $color;
                $result->MinorValueGridlineWidthPt = $widthPt;
            }
        } else {
            $result->HasValueGridlines = $hasVisibleLine || $line === null;
            if ($hasVisibleLine) {
                $result->ValueGridlineColor = $color;
                $result->ValueGridlineWidthPt = $widthPt;
            }
        }
    }

    public static function ParseLineStyle($line, $colors, &$hasVisibleLine, &$color, &$widthPt)
    {
        $hasVisibleLine = false;
        $color = new SKColor(0xD9, 0xD9, 0xD9);
        $widthPt = 1.0;

        if ($line === null) {
            return;
        }

        if ($line->Element(self::A . 'noFill') !== null) {
            return;
        }

        $emu = 0.0;
        if (self::TryDouble(self::AttrOf($line, 'w'), $emu) && $emu > 0.0) {
            $widthPt = $emu / 12700.0;
        }

        $solidFill = $line->Element(self::A . 'solidFill');
        $fill = self::FillChild($solidFill);
        if ($fill !== null) {
            $parsedColor = null;
            if (self::TryDrawingColor($fill, $colors, $parsedColor)) {
                $color = $parsedColor;
                $hasVisibleLine = true;
                return;
            }
        }

        $hasVisibleLine = $line->Element(self::A . 'solidFill') === null;
    }

    public static function TryDrawingColor($source, $colors, &$color)
    {
        $color = SKColors::Transparent();
        if ($source === null || $colors === null) {
            return false;
        }

        if ($source->Name === self::A . 'srgbClr') {
            return self::TryHexWithAlpha(self::AttrOf($source, 'val'), $source, $color);
        }

        if ($source->Name === self::A . 'schemeClr') {
            $baseColor = $colors->ResolveSchemeName(self::AttrOf($source, 'val'), SKColors::Transparent());
            $lumMod = 0.0;
            $lumOff = 0.0;
            self::TryDouble(self::AttrOf($source->Element(self::A . 'lumMod'), 'val'), $lumMod);
            self::TryDouble(self::AttrOf($source->Element(self::A . 'lumOff'), 'val'), $lumOff);
            if ($lumMod > 0.0 || $lumOff > 0.0) {
                $baseColor = RenderColor::ApplyLuma($baseColor, $lumMod > 0.0 ? $lumMod / 100000.0 : 1.0, $lumOff / 100000.0);
            }

            $alpha = 0.0;
            if (self::TryDouble(self::AttrOf($source->Element(self::A . 'alpha'), 'val'), $alpha) && $alpha >= 0.0) {
                $channel = (int)round(min(100000.0, $alpha) / 100000.0 * 255.0, 0, PHP_ROUND_HALF_EVEN);
                $color = new SKColor($baseColor->Red, $baseColor->Green, $baseColor->Blue, $channel);
            } else {
                $color = $baseColor;
            }

            return true;
        }

        return false;
    }

    private static function TryHexWithAlpha($hex, $source, &$color)
    {
        $color = SKColors::Transparent();
        $rgb = null;
        if (!self::TryHex($hex, $rgb)) {
            return false;
        }

        $alpha = 0.0;
        if (self::TryDouble(self::AttrOf($source->Element(self::A . 'alpha'), 'val'), $alpha) && $alpha >= 0.0) {
            $channel = (int)round(min(100000.0, $alpha) / 100000.0 * 255.0, 0, PHP_ROUND_HALF_EVEN);
            $color = new SKColor($rgb->Red, $rgb->Green, $rgb->Blue, $channel);
        } else {
            $color = $rgb;
        }

        return true;
    }

    private static function ParseSeriesColor($ser, $colors, $index)
    {
        $line = $ser->Element(self::C . 'spPr') !== null ? $ser->Element(self::C . 'spPr')->Element(self::A . 'ln') : null;
        $fill = $ser->Element(self::C . 'spPr') !== null ? $ser->Element(self::C . 'spPr')->Element(self::A . 'solidFill') : null;
        $source = self::FirstNonNull(self::FillChild($fill), self::FillChild($line !== null ? $line->Element(self::A . 'solidFill') : null));

        if ($source !== null) {
            if ($source->Name === self::A . 'srgbClr') {
                $parsed = null;
                if (self::TryHex(self::AttrOf($source, 'val'), $parsed)) return $parsed;
            } elseif ($source->Name === self::A . 'schemeClr') {
                $name = self::AttrOf($source, 'val');
                $baseColor = $colors->ResolveSchemeName($name, self::DefaultSeriesColor($colors, $index));
                $lumMod = 0.0;
                $lumOff = 0.0;
                self::TryDouble(self::AttrOf($source->Element(self::A . 'lumMod'), 'val'), $lumMod);
                self::TryDouble(self::AttrOf($source->Element(self::A . 'lumOff'), 'val'), $lumOff);
                if ($lumMod > 0.0 || $lumOff > 0.0) {
                    return RenderColor::ApplyLuma($baseColor, $lumMod > 0.0 ? $lumMod / 100000.0 : 1.0, $lumOff / 100000.0);
                }
                return $baseColor;
            }
        }

        return self::DefaultSeriesColor($colors, $index);
    }

    private static function ParseSeriesPatternFill($ser, $colors, $series)
    {
        $shapeProperties = $ser !== null ? $ser->Element(self::C . 'spPr') : null;
        $patternFill = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'pattFill') : null;
        if ($patternFill === null || $series === null) {
            return;
        }

        $series->PatternPreset = self::AttrOf($patternFill, 'prst');
        $parsedColor = null;
        $foreground = self::FillChild($patternFill->Element(self::A . 'fgClr'));
        if (self::TryDrawingColor($foreground, $colors, $parsedColor)) {
            $series->PatternForegroundColor = $parsedColor;
            $series->Color = $parsedColor;
        }

        $background = self::FillChild($patternFill->Element(self::A . 'bgClr'));
        if (self::TryDrawingColor($background, $colors, $parsedColor)) {
            $series->PatternBackgroundColor = $parsedColor;
        }
    }

    public static function FillChild($solidFill)
    {
        if ($solidFill === null) return null;
        return $solidFill->Element(self::A . 'srgbClr') ?? $solidFill->Element(self::A . 'schemeClr');
    }

    private static function ParsePiePointColors($ser, $colors, $pointCount)
    {
        $result = array();
        if ($pointCount < 0) {
            $pointCount = 0;
        }

        for ($i = 0; $i < $pointCount + 1; $i++) {
            $result[] = self::DefaultSeriesColor($colors, $i);
        }

        if ($ser === null) {
            return $result;
        }

        foreach ($ser->Elements(self::C . 'dPt') as $point) {
            $index = -1;
            $indexText = self::AttrOf($point->Element(self::C . 'idx'), 'val');
            if ($indexText === null || !is_numeric($indexText)) {
                continue;
            }

            $index = (int)$indexText;
            while ($index >= count($result)) {
                $result[] = self::DefaultSeriesColor($colors, count($result));
            }

            $shapeProperties = $point->Element(self::C . 'spPr');
            $fill = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'solidFill') : null;
            $source = self::FillChild($fill);
            $color = null;
            if (self::TryDrawingColor($source, $colors, $color)) {
                $result[$index] = $color;
            }
        }

        return $result;
    }

    private static function TryResolveTextColor($textProperties, $colors, $chartModel, &$color)
    {
        $color = SKColors::Transparent();
        $runProperties = self::FirstDescendant($textProperties, self::A . 'rPr');
        if (ChartTextStyleResolver::TryResolveTextColor($runProperties, $colors, $chartModel, $color)) {
            return true;
        }

        $defaultRunProperties = self::ResolveDefaultRunProperties($textProperties);
        if ($defaultRunProperties === null) {
            return false;
        }

        if (ChartTextStyleResolver::TryResolveTextColor($defaultRunProperties, $colors, $chartModel, $color)) {
            return true;
        }

        $paragraph = $textProperties->Element(self::A . 'p');
        $endParagraphRunProperties = $paragraph !== null ? $paragraph->Element(self::A . 'endParaRPr') : null;
        return ChartTextStyleResolver::TryResolveTextColor($endParagraphRunProperties, $colors, $chartModel, $color);
    }

    private static function TryResolveTextOutlineColor($textProperties, $colors, &$color)
    {
        $color = SKColors::Transparent();
        $runProperties = self::FirstDescendant($textProperties, self::A . 'rPr');
        if (ChartTextStyleResolver::TryResolveTextOutlineColor($runProperties, $colors, $color)) {
            return true;
        }

        $defaultRunProperties = self::ResolveDefaultRunProperties($textProperties);
        return ChartTextStyleResolver::TryResolveTextOutlineColor($defaultRunProperties, $colors, $color);
    }

    private static function TryResolveTextFont($textProperties, $colors, &$font)
    {
        $font = null;
        $defaultRunProperties = self::ResolveDefaultRunProperties($textProperties);
        if ($defaultRunProperties === null) {
            return false;
        }

        return ChartTextStyleResolver::TryResolveFont($defaultRunProperties,
            $colors->MajorFontName, $colors->MinorFontName, $font);
    }

    private static function ResolveTextProperties($owner)
    {
        if ($owner === null) {
            return null;
        }

        $richText = $owner->Element(self::C . 'tx');
        if ($richText !== null) {
            $richText = $richText->Element(self::C . 'rich');
        }

        return $richText ?? $owner->Element(self::C . 'txPr');
    }

    private static function ResolveDefaultRunProperties($textProperties)
    {
        if ($textProperties === null) {
            return null;
        }

        $paragraph = $textProperties->Element(self::A . 'p');
        $paragraphProperties = $paragraph !== null ? $paragraph->Element(self::A . 'pPr') : null;
        return $paragraphProperties !== null ? $paragraphProperties->Element(self::A . 'defRPr') : null;
    }

    private static function DefaultSeriesColor($colors, $index)
    {
        // Cycle the Office accent slots when a series has no explicit color.
        return $colors->ResolveSchemeName('accent' . (($index % 6) + 1), SKColors::Gray());
    }

    private static function ParseLineWidthPt($ser)
    {
        $line = $ser->Element(self::C . 'spPr') !== null ? $ser->Element(self::C . 'spPr')->Element(self::A . 'ln') : null;
        $emu = 0.0;
        if ($line !== null && self::TryDouble(self::AttrOf($line, 'w'), $emu) && $emu > 0.0) {
            return $emu / 12700.0;
        }
        return 2.25;
    }

    private static function ParseSeriesLineStyle($ser, $colors, $series)
    {
        if ($ser === null || $colors === null || $series === null) {
            return;
        }

        $series->LineColor = $series->Color;
        $series->HasVisibleLine = true;

        $line = $ser->Element(self::C . 'spPr') !== null ? $ser->Element(self::C . 'spPr')->Element(self::A . 'ln') : null;
        if ($line === null) {
            return;
        }

        if ($line->Element(self::A . 'noFill') !== null) {
            $series->HasVisibleLine = false;
            return;
        }

        $color = null;
        $widthPt = 0.0;
        $hasVisibleLine = false;
        self::ParseLineStyle($line, $colors, $hasVisibleLine, $color, $widthPt);
        $series->HasVisibleLine = $hasVisibleLine;
        if ($hasVisibleLine) {
            $series->LineColor = $color;
            $series->LineWidthPt = $widthPt;
        }
    }

    private static function ParseMarker($ser, $colors, $series)
    {
        if ($ser === null || $series === null) {
            return;
        }

        $series->MarkerFillColor = $series->Color;
        $series->MarkerStrokeColor = $series->Color;
        $series->MarkerStrokeWidthPt = 0.75;

        $marker = $ser->Element(self::C . 'marker');
        if ($marker === null) {
            return;
        }

        $symbol = self::AttrOf($marker->Element(self::C . 'symbol'), 'val');
        if ($symbol !== null && $symbol !== '') {
            $series->MarkerSymbol = $symbol;
            $series->MarkerVisible = $symbol !== 'none';
        }

        $size = 0;
        $sizeText = self::AttrOf($marker->Element(self::C . 'size'), 'val');
        if ($sizeText !== null && is_numeric($sizeText)) {
            $size = (int)$sizeText;
        }
        if ($size > 0) {
            $series->MarkerSize = $size;
        }

        $markerShape = $marker->Element(self::C . 'spPr');
        $fill = $markerShape !== null ? $markerShape->Element(self::A . 'solidFill') : null;
        $fillSource = self::FillChild($fill);
        $markerFillColor = null;
        if (self::TryDrawingColor($fillSource, $colors, $markerFillColor)) {
            $series->MarkerFillColor = $markerFillColor;
        }

        $line = $markerShape !== null ? $markerShape->Element(self::A . 'ln') : null;
        if ($line !== null) {
            $markerStrokeColor = null;
            $markerStrokeWidth = 0.0;
            $hasVisibleLine = false;
            self::ParseLineStyle($line, $colors, $hasVisibleLine, $markerStrokeColor, $markerStrokeWidth);
            if ($hasVisibleLine) {
                $series->MarkerStrokeColor = $markerStrokeColor;
                $series->MarkerStrokeWidthPt = $markerStrokeWidth;
            }
        }
    }

    private static function ApplyRadarSeriesDefaults($chart, $series)
    {
        if ($chart === null || $series === null || $chart->Kind !== ChartKind::Radar) {
            return;
        }
    }

    private static function ParseNumericData($holder, $workbook, $dateSystem)
    {
        $formula = self::TextOf(self::FirstDescendant($holder, self::C . 'f'));
        $resolved = ChartWorkbookDataResolver::ResolveNumericRange($workbook, $formula, $dateSystem);
        if (count($resolved) > 0) {
            return $resolved;
        }

        return self::ParseNumericCache($holder);
    }

    private static function ParseNumericCache($holder)
    {
        $result = array();
        if ($holder === null) {
            return $result;
        }

        $cache = self::Descendant($holder, self::C . 'numCache') ?? self::Descendant($holder, self::C . 'numLit');
        if ($cache === null) {
            return $result;
        }

        $count = 0;
        $countText = self::AttrOf($cache->Element(self::C . 'ptCount'), 'val');
        if ($countText !== null && is_numeric($countText)) {
            $count = (int)$countText;
        }
        $values = array_fill(0, max(0, $count), null);
        foreach ($cache->Elements(self::C . 'pt') as $pt) {
            $idx = -1;
            $v = 0.0;
            $idxText = self::AttrOf($pt, 'idx');
            if ($idxText !== null && is_numeric($idxText)) {
                $idx = (int)$idxText;
            }
            if ($idx >= 0 && $idx < count($values)
                && self::TryDouble(self::TextOf($pt->Element(self::C . 'v')), $v)) {
                $values[$idx] = $v;
            }
        }
        foreach ($values as $value) {
            $result[] = $value;
        }
        return $result;
    }

    private static function ParseCategories($catEl, $workbook, $dateSystem, $culture)
    {
        $labels = array();
        if ($catEl === null) {
            return $labels;
        }

        $stringFormula = self::TextOf(self::FirstDescendant($catEl->Element(self::C . 'strRef'), self::C . 'f'));
        if ($stringFormula !== null && $stringFormula !== '') {
            $resolvedStringLabels = ChartWorkbookDataResolver::ResolveStringRange($workbook, $stringFormula);
            if (count($resolvedStringLabels) > 0) {
                return $resolvedStringLabels;
            }
        }

        $numericFormula = self::TextOf(self::FirstDescendant($catEl->Element(self::C . 'numRef'), self::C . 'f'));
        $numericFormat = self::TextOf(self::FirstDescendant($catEl, self::C . 'formatCode'));
        $resolvedLabels = ChartWorkbookDataResolver::ResolveCategoryLabels($workbook, $numericFormula, $numericFormat, $dateSystem, $culture);
        if (count($resolvedLabels) > 0) {
            return $resolvedLabels;
        }

        $strCache = self::Descendant($catEl, self::C . 'strCache') ?? self::Descendant($catEl, self::C . 'strLit');
        if ($strCache !== null) {
            $count = 0;
            $countText = self::AttrOf($strCache->Element(self::C . 'ptCount'), 'val');
            if ($countText !== null && is_numeric($countText)) {
                $count = (int)$countText;
            }
            $arr = array_fill(0, max(0, $count), null);
            foreach ($strCache->Elements(self::C . 'pt') as $pt) {
                $idx = -1;
                $idxText = self::AttrOf($pt, 'idx');
                if ($idxText !== null && is_numeric($idxText)) {
                    $idx = (int)$idxText;
                }
                if ($idx >= 0 && $idx < count($arr)) {
                    $arr[$idx] = self::TextOf($pt->Element(self::C . 'v'));
                }
            }
            for ($i = 0; $i < count($arr); $i++) $labels[] = $arr[$i] ?? '';
            return $labels;
        }

        $numCache = self::Descendant($catEl, self::C . 'numCache');
        if ($numCache !== null) {
            $code = self::TextOf($numCache->Element(self::C . 'formatCode'));
            $count = 0;
            $countText = self::AttrOf($numCache->Element(self::C . 'ptCount'), 'val');
            if ($countText !== null && is_numeric($countText)) {
                $count = (int)$countText;
            }
            $arr = array_fill(0, max(0, $count), null);
            foreach ($numCache->Elements(self::C . 'pt') as $pt) {
                $idx = -1;
                $v = 0.0;
                $idxText = self::AttrOf($pt, 'idx');
                if ($idxText !== null && is_numeric($idxText)) {
                    $idx = (int)$idxText;
                }
                if ($idx >= 0 && $idx < count($arr)
                    && self::TryDouble(self::TextOf($pt->Element(self::C . 'v')), $v)) {
                    $arr[$idx] = self::FormatLabel($v, $code, $dateSystem, $culture);
                }
            }
            for ($i = 0; $i < count($arr); $i++) $labels[] = $arr[$i] ?? '';
        }

        return $labels;
    }

    private static function ParseCategoryValues($catEl, $workbook, $dateSystem)
    {
        return self::ParseNumericData($catEl, $workbook, $dateSystem);
    }

    // Formats a numeric axis/category value with its cached format code.
    public static function FormatLabel($value, $formatCode, $dateSystem, $culture)
    {
        if ($formatCode === null || $formatCode === '' || $formatCode === 'General') {
            return sprintf('%.15G', $value);
        }

        $style = new StyleValue();
        $style->NumberFormat = new NumberFormatValue();
        $style->NumberFormat->Custom = $formatCode;
        $boxed = $value;
        if (self::IsDateFormat($formatCode)) {
            $boxed = DateSerialConverter::FromSerial($value, $dateSystem);
        }

        try {
            return DisplayTextFormatter::FormatDisplayValue($boxed, $style, $culture !== null ? $culture : CultureInfo::GetInvariantCulture());
        } catch (\Throwable $exception) {
            return sprintf('%.15G', $value);
        }
    }

    public static function IsDateFormat($code)
    {
        if ($code === null || $code === '') {
            return false;
        }

        // Ignore color/condition blocks like [Red] (whose letters would otherwise look like date
        // tokens), backslash escapes, and quoted literals before checking for real date tokens.
        $cleaned = preg_replace('/\[[^\]]*\]/', '', $code);
        $cleaned = preg_replace('/\\\\./', '', $cleaned);
        $cleaned = preg_replace('/"[^"]*"/', '', $cleaned);

        for ($i = 0; $i < strlen($cleaned); $i++) {
            $ch = $cleaned[$i];
            if ($ch === 'y' || $ch === 'd' || $ch === 'h' || $ch === 's') {
                return true;
            }
        }

        return false;
    }

    // --- small XML helpers ---

    private static function ParseTitleText($titleEl, $workbook)
    {
        $textSource = $titleEl->Element(self::C . 'tx');
        $titleFormula = self::TextOf(self::FirstDescendant($textSource !== null ? $textSource->Element(self::C . 'strRef') : null, self::C . 'f'));
        if ($titleFormula !== null && $titleFormula !== '') {
            $resolvedTitle = ChartWorkbookDataResolver::ResolveSingleString($workbook, $titleFormula);
            if ($resolvedTitle !== null && $resolvedTitle !== '') {
                return $resolvedTitle;
            }
        }

        $cachedValue = self::FirstDescendant($textSource, self::C . 'v');
        if ($cachedValue !== null) {
            $cachedText = self::TextOf($cachedValue);
            if ($cachedText !== null && $cachedText !== '') {
                return $cachedText;
            }
        }

        return self::ConcatText($titleEl);
    }

    private static function ParseAxisTitle($axisElement, $workbook)
    {
        $title = $axisElement !== null ? $axisElement->Element(self::C . 'title') : null;
        if ($title === null) {
            return null;
        }

        $text = self::ParseTitleText($title, $workbook);
        return $text !== null && $text !== '' ? trim($text) : 'Axis Title';
    }

    private static function ParseSeriesName($textSource, $workbook)
    {
        $nameFormula = self::TextOf(self::FirstDescendant($textSource !== null ? $textSource->Element(self::C . 'strRef') : null, self::C . 'f'));
        if ($nameFormula !== null && $nameFormula !== '') {
            $resolvedName = ChartWorkbookDataResolver::ResolveSingleString($workbook, $nameFormula);
            if ($resolvedName !== null && $resolvedName !== '') {
                return $resolvedName;
            }
        }

        $cachedValue = self::FirstDescendant($textSource, self::C . 'v');
        if ($cachedValue !== null) {
            $cachedText = self::TextOf($cachedValue);
            if ($cachedText !== null && $cachedText !== '') {
                return $cachedText;
            }
        }

        return self::ConcatText($textSource);
    }

    private static function ConcatText($element)
    {
        if ($element === null) {
            return '';
        }

        $sb = '';
        foreach ($element->Descendants(self::A . 't') as $t) {
            $sb .= $t->Value;
        }
        return $sb;
    }

    private static function FirstElement($parent, $name)
    {
        foreach ($parent->Elements($name) as $e) return $e;
        return null;
    }

    private static function FirstDescendant($parent, $name)
    {
        if ($parent === null) return null;
        foreach ($parent->Descendants($name) as $e) return $e;
        return null;
    }

    private static function Descendant($parent, $name)
    {
        return self::FirstDescendant($parent, $name);
    }

    private static function FirstNonNull($a, $b)
    {
        return $a ?? $b;
    }

    private static function TextOf($e)
    {
        return $e !== null ? $e->Value : '';
    }

    private static function AttrOf($e, $name)
    {
        if ($e === null) return null;
        $a = $e->Attribute($name);
        return $a !== null ? $a->Value : null;
    }

    private static function TryDouble($s, &$value)
    {
        $value = 0.0;
        if ($s === null || !is_numeric($s)) {
            return false;
        }

        $value = (float)$s;
        return true;
    }

    private static function TryHex($hex, &$color)
    {
        $color = SKColors::Gray();
        if ($hex === null || $hex === '' || strlen($hex) < 6) return false;
        $offset = strlen($hex) === 8 ? 2 : 0;
        $rText = substr($hex, $offset, 2);
        $gText = substr($hex, $offset + 2, 2);
        $bText = substr($hex, $offset + 4, 2);
        if (strlen($rText) === 2 && strlen($gText) === 2 && strlen($bText) === 2
            && ctype_xdigit($rText) && ctype_xdigit($gText) && ctype_xdigit($bText)) {
            $color = new SKColor(hexdec($rText), hexdec($gText), hexdec($bText));
            return true;
        }
        return false;
    }
}
