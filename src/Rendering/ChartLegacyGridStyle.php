<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;

final class ChartLegacyGridStyle
{
    const C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}';

    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    const Mc = '{http://schemas.openxmlformats.org/markup-compatibility/2006}';

    private function __construct()
    {
    }

    public static function Apply($root, $plot, $model, $chart)
    {
        if ($root->Element(self::Mc . 'AlternateContent') !== null) {
            return;
        }

        if ($model !== null) {
            foreach ($model->CompanionFiles as $companion) {
                if ($companion->RelationshipType === 'http://schemas.microsoft.com/office/2011/relationships/chartStyle') {
                    return;
                }
            }
        }

        $styleElement = $root->Element(self::C . 'style');
        $styleAttribute = $styleElement !== null ? $styleElement->Attribute('val') : null;
        $style = $styleAttribute !== null ? (string)$styleAttribute : null;
        if ($style !== null && $style !== '2' && $style !== '3') {
            return;
        }

        $axis = $plot->Element(self::C . 'valAx');
        $grid = $axis !== null ? $axis->Element(self::C . 'majorGridlines') : null;
        $gridSpPr = $grid !== null ? $grid->Element(self::C . 'spPr') : null;
        $gridLine = $gridSpPr !== null ? $gridSpPr->Element(self::A . 'ln') : null;
        if ($grid !== null) {
            if ($gridLine === null || $gridLine->Attribute('w') === null) {
                $chart->ValueGridlineWidthPt = 0.75;
            }

            if ($gridLine === null) {
                $chart->ValueGridlineColor = new SKColor(0x86, 0x86, 0x86);
            }
        }

        if ($axis === null) {
            return;
        }

        $deleteElement = $axis->Element(self::C . 'delete');
        $deleteAttribute = $deleteElement !== null ? $deleteElement->Attribute('val') : null;
        $deleted = $deleteAttribute !== null ? (string)$deleteAttribute : null;
        if ($deleted === '1' || $deleted === 'true') {
            $chart->HasValueAxisLine = false;
            return;
        }

        // Missing line properties inherit the legacy style; explicit noFill does not.
        $axisSpPr = $axis->Element(self::C . 'spPr');
        $axisLine = $axisSpPr !== null ? $axisSpPr->Element(self::A . 'ln') : null;
        if ($axisLine === null) {
            $chart->HasValueAxisLine = true;
            $chart->ValueAxisLineColor = new SKColor(0x86, 0x86, 0x86);
            $chart->ValueAxisLineWidthPt = 0.75;
        }

        if ($chart->ValueAxisMajorTickMark === null) {
            $chart->ValueAxisMajorTickMark = 'out';
        }
    }
}
