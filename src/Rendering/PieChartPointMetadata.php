<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColors;

final class PieChartPointMetadata
{
    const C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}';

    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function Apply($series, $colors, $chart)
    {
        if ($chart === null || $chart->Kind !== ChartKind::Pie || count($chart->Series) === 0) {
            return;
        }

        $pointCount = count($chart->Series[0]->Values);
        if (count($chart->Categories) === 0) {
            for ($i = 0; $i < $pointCount; $i++) {
                $chart->Categories[] = (string)($i + 1);
            }
        }

        for ($i = 0; $i < $pointCount; $i++) {
            $chart->PiePointLineColors[] = SKColors::Transparent();
            $chart->PiePointLineWidthsPt[] = 0.0;
        }

        if ($series === null) {
            return;
        }

        foreach ($series->Elements(self::C . 'dPt') as $point) {
            $indexElement = $point->Element(self::C . 'idx');
            if ($indexElement === null) {
                continue;
            }

            $indexText = (string)$indexElement->Attribute('val');
            $index = -1;
            if ($indexText !== '' && is_numeric($indexText) && (int)$indexText == (float)$indexText) {
                $index = (int)$indexText;
            }

            if ($index < 0 || $index >= $pointCount) {
                continue;
            }

            $shape = $point->Element(self::C . 'spPr');
            $line = $shape !== null ? $shape->Element(self::A . 'ln') : null;
            ChartXmlParser::ParseLineStyle($line, $colors, $visible, $color, $width);
            if ($visible) {
                $chart->PiePointLineColors[$index] = $color;
                $chart->PiePointLineWidthsPt[$index] = $width;
            }
        }
    }
}
