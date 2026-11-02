<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Skia\SKColors;

final class ChartExplicitFormatting
{
    const C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}';

    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function ApplyTextDefaults($root, $colors, $chart)
    {
        $txPr = $root->Element(self::C . 'txPr');
        $p = $txPr !== null ? $txPr->Element(self::A . 'p') : null;
        $pPr = $p !== null ? $p->Element(self::A . 'pPr') : null;
        $properties = $pPr !== null ? $pPr->Element(self::A . 'defRPr') : null;
        $size = self::ReadSize($properties);
        if ($size === null) {
            return;
        }

        $chart->TextSizePt = $size;
        $chart->TitleFont = new FontValue();
        $chart->TitleFont->Name = $colors->MinorFontName !== null ? $colors->MinorFontName : 'Calibri';
        $chart->TitleFont->Size = $size * 1.2;
        $chart->TitleFont->Bold = true;
        $color = $colors->ResolveSchemeName('tx1', SKColors::Black());
        if (ChartXmlParser::TryDrawingColor(ChartXmlParser::FillChild($properties->Element(self::A . 'solidFill')), $colors, $explicitColor)) {
            $color = $explicitColor;
        }

        $chart->TitleColor = $color;
        $chart->AxisLabelColor = $color;
        $chart->LegendTextColor = $color;
    }

    public static function ApplySeries($element, $type, $colors, $series, $fallbackIndex)
    {
        $idxElement = $element->Element(self::C . 'idx');
        $idxAttribute = $idxElement !== null ? $idxElement->Attribute('val') : null;
        $index = -1;
        $text = $idxAttribute !== null ? (string)$idxAttribute : '';
        if ($text !== '' && is_numeric($text) && (int)$text == (float)$text) {
            $index = (int)$text;
        }
        if ($index < 0) {
            $index = $fallbackIndex;
        }

        $shape = $element->Element(self::C . 'spPr');
        $shapeSolidFill = $shape !== null ? $shape->Element(self::A . 'solidFill') : null;
        $shapeLine = $shape !== null ? $shape->Element(self::A . 'ln') : null;
        $lineSolidFill = $shapeLine !== null ? $shapeLine->Element(self::A . 'solidFill') : null;
        if ($shapeSolidFill === null && $lineSolidFill === null) {
            $series->Color = self::AutomaticColor($colors, $index);
        }

        $varyColors = $type->Element(self::C . 'varyColors');
        $varyColorsAttribute = $varyColors !== null ? $varyColors->Attribute('val') : null;
        if (($varyColorsAttribute !== null ? (string)$varyColorsAttribute : '') === '1') {
            for ($i = 0; $i < count($series->Values); $i++) {
                $series->PointColors[$i] = self::AutomaticColor($colors, $i);
            }
        }

        foreach ($element->Elements(self::C . 'dPt') as $point) {
            $idxAttribute = null;
            $idxElement = $point->Element(self::C . 'idx');
            if ($idxElement !== null) {
                $idxAttribute = $idxElement->Attribute('val');
            }

            $pointText = $idxAttribute !== null ? (string)$idxAttribute : '';
            $pointIndex = -1;
            if ($pointText !== '' && is_numeric($pointText) && (int)$pointText == (float)$pointText) {
                $pointIndex = (int)$pointText;
            }

            $spPr = $point->Element(self::C . 'spPr');
            $solidFill = $spPr !== null ? $spPr->Element(self::A . 'solidFill') : null;
            if ($pointIndex >= 0
                && ChartXmlParser::TryDrawingColor(ChartXmlParser::FillChild($solidFill), $colors, $color)) {
                $series->PointColors[$pointIndex] = $color;
            }
        }

        $labels = $element->Element(self::C . 'dLbls');
        if ($labels === null) {
            $labels = $type->Element(self::C . 'dLbls');
        }

        $txPr = $labels !== null ? $labels->Element(self::C . 'txPr') : null;
        $p = $txPr !== null ? $txPr->Element(self::A . 'p') : null;
        $pPr = $p !== null ? $p->Element(self::A . 'pPr') : null;
        $defRPr = $pPr !== null ? $pPr->Element(self::A . 'defRPr') : null;
        $series->DataLabelSizePt = self::ReadSize($defRPr);
    }

    public static function PointColor($series, $index)
    {
        return array_key_exists($index, $series->PointColors) ? $series->PointColors[$index] : $series->Color;
    }

    private static function AutomaticColor($colors, $index)
    {
        $color = $colors->ResolveSchemeName('accent' . ($index % 6 + 1), SKColors::Gray());
        if ($index >= 6) {
            $color = RenderColor::ApplyLuma($color, 0.6, 0.4);
        }

        return $color;
    }

    private static function ReadSize($properties)
    {
        $attribute = $properties !== null ? $properties->Attribute('sz') : null;
        $text = $attribute !== null ? (string)$attribute : '';
        if (is_numeric($text) && (float)$text > 0.0) {
            return (float)$text / 100.0;
        }

        return null;
    }
}
