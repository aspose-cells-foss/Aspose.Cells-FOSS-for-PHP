<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class MapChartXmlParser
{
    const Cx = '{http://schemas.microsoft.com/office/drawing/2014/chartex}';

    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function Parse($xml, $workbook)
    {
        if ($xml === null || $xml === '' || $workbook === null) {
            return null;
        }

        $document = null;
        try {
            $document = XDocument::Parse($xml);
        } catch (\Exception $exception) {
            return null;
        }

        $chart = $document->Root !== null ? $document->Root->Element(self::Cx . 'chart') : null;
        $region = $chart !== null ? $chart->Element(self::Cx . 'plotArea') : null;
        $region = $region !== null ? $region->Element(self::Cx . 'plotAreaRegion') : null;
        if ($region === null) {
            return null;
        }

        $visibleSeries = null;
        foreach ($region->Elements(self::Cx . 'series') as $series) {
            if (self::AttrOf($series, 'layoutId') === 'regionMap'
                && self::AttrOf($series, 'hidden') !== '1') {
                $visibleSeries = $series;
                break;
            }
        }

        if ($visibleSeries === null) {
            return null;
        }

        $dataId = $visibleSeries->Element(self::Cx . 'dataId');
        $chartData = $document->Root->Element(self::Cx . 'chartData');
        $data = null;
        if ($dataId !== null && $chartData !== null) {
            foreach ($chartData->Elements(self::Cx . 'data') as $candidate) {
                if (self::AttrOf($candidate, 'id') === self::AttrOf($dataId, 'val')) {
                    $data = $candidate;
                    break;
                }
            }
        }

        if ($data === null) {
            return null;
        }

        $categoryDimension = $data->Element(self::Cx . 'strDim');
        $numericDimension = $data->Element(self::Cx . 'numDim');
        $categoryName = $categoryDimension !== null ? self::TextOf($categoryDimension->Element(self::Cx . 'f')) : null;
        $numericName = $numericDimension !== null ? self::TextOf($numericDimension->Element(self::Cx . 'f')) : null;
        $categories = ChartWorkbookDataResolver::ResolveStringRange($workbook, self::ResolveName($workbook, $categoryName));
        $values = ChartWorkbookDataResolver::ResolveNumericRange($workbook, self::ResolveName($workbook, $numericName), $workbook->Settings->DateSystem);
        $displayValues = ChartWorkbookDataResolver::ResolveStringRange($workbook, self::ResolveName($workbook, $numericName));
        if (count($categories) === 0 || count($categories) !== count($values)) {
            return null;
        }

        $result = new MapChartData();
        $title = $chart->Element(self::Cx . 'title');
        $titleData = $title !== null ? $title->Element(self::Cx . 'tx') : null;
        $titleData = $titleData !== null ? $titleData->Element(self::Cx . 'txData') : null;
        $titleValue = $titleData !== null ? $titleData->Element(self::Cx . 'v') : null;
        $titleText = $titleValue !== null ? $titleValue->Value : null;
        $result->Title = $title !== null
            ? ($titleText !== null ? $titleText : 'Chart Title')
            : '';
        $seriesTitle = $visibleSeries->Element(self::Cx . 'tx');
        $seriesTitle = $seriesTitle !== null ? $seriesTitle->Element(self::Cx . 'txData') : null;
        $result->SeriesName = $seriesTitle !== null ? self::TextOf($seriesTitle->Element(self::Cx . 'v')) : '';
        self::ParseColors($visibleSeries->Element(self::Cx . 'valueColors'), RenderColor::FromWorkbook($workbook), $result);

        $hasValues = false;
        for ($i = 0; $i < count($categories); $i++) {
            if ($categories[$i] === null || $categories[$i] === '' || $values[$i] === null) {
                continue;
            }

            $value = $values[$i];
            $result->Values[$categories[$i]] = $value;
            if (!$hasValues) {
                $result->Minimum = $value;
                $result->Maximum = $value;
                $result->MinimumLabel = $i < count($displayValues) ? $displayValues[$i] : '';
                $result->MaximumLabel = $result->MinimumLabel;
                $hasValues = true;
            } else {
                if ($value < $result->Minimum) {
                    $result->Minimum = $value;
                    $result->MinimumLabel = $i < count($displayValues) ? $displayValues[$i] : '';
                }
                if ($value > $result->Maximum) {
                    $result->Maximum = $value;
                    $result->MaximumLabel = $i < count($displayValues) ? $displayValues[$i] : '';
                }
            }
        }

        return $hasValues ? $result : null;
    }

    private static function ParseColors($valueColors, $colors, $result)
    {
        if ($valueColors === null) {
            return;
        }

        $minimum = null;
        $maximum = null;
        if (!self::TryColor($valueColors->Element(self::Cx . 'minColor'), $colors, $minimum)
            || !self::TryColor($valueColors->Element(self::Cx . 'maxColor'), $colors, $maximum)) {
            return;
        }

        $result->HasCustomColors = true;
        $result->MinimumColor = $minimum;
        $result->MaximumColor = $maximum;
        $middle = null;
        if (self::TryColor($valueColors->Element(self::Cx . 'midColor'), $colors, $middle)) {
            $result->HasMiddleColor = true;
            $result->MiddleColor = $middle;
        }
    }

    private static function TryColor($container, $colors, &$color)
    {
        $color = SKColors::Transparent();
        if ($container === null) {
            return false;
        }

        $scheme = $container->Element(self::A . 'schemeClr');
        if ($scheme !== null) {
            $color = $colors->ResolveSchemeName(self::AttrOf($scheme, 'val'), SKColors::Transparent());
            return $color->Alpha > 0;
        }

        $rgb = $container->Element(self::A . 'srgbClr');
        $text = $rgb !== null ? self::AttrOf($rgb, 'val') : null;
        if ($rgb !== null && $text !== null && $text !== '' && ctype_xdigit($text)) {
            $value = hexdec($text);
            $color = new SKColor(((int)$value >> 16) & 0xFF, ((int)$value >> 8) & 0xFF, (int)$value & 0xFF);
            return true;
        }

        return false;
    }

    private static function ResolveName($workbook, $formula)
    {
        if ($formula === null || $formula === '') {
            return '';
        }

        for ($i = 0; $i < count($workbook->DefinedNames); $i++) {
            $name = $workbook->DefinedNames[$i];
            if ($name->Name === $formula) {
                return $name->Formula;
            }
        }

        return $formula;
    }

    private static function TextOf($e)
    {
        return $e !== null ? $e->Value : '';
    }

    private static function AttrOf($e, $name)
    {
        if ($e === null) {
            return null;
        }
        $a = $e->Attribute($name);
        return $a !== null ? $a->Value : null;
    }
}
