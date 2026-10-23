<?php

namespace Aspose\Cells_FOSS\Samples\Charts;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\ChartType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'charts-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Charts';

        $sheet->Cells['A1']->PutValue('Month');
        $sheet->Cells['B1']->PutValue('Sales');
        $sheet->Cells['C1']->PutValue('Profit');

        for ($month = 1; $month <= 12; $month++) {
            $sheet->Cells[[$month, 0]]->PutValue('Month ' . $month);
            $sheet->Cells[[$month, 1]]->PutValue($month * 1000 + $month * 50);
            $sheet->Cells[[$month, 2]]->PutValue($month * 300 + $month * 20);
        }

        $salesChartIndex = $sheet->Charts->Add(ChartType::Column, 'Charts!$B$1:$B$13', 0, 4, 18, 8);
        $salesChart = $sheet->Charts[$salesChartIndex];
        $sheet->Cells['E1']->PutValue('Sales Chart: ' . $salesChart->Name);

        $profitChartIndex = $sheet->Charts->Add(ChartType::Line, 'Charts!$C$1:$C$13', 0, 9, 18, 13);
        $profitChart = $sheet->Charts[$profitChartIndex];
        $sheet->Cells['J1']->PutValue('Profit Chart: ' . $profitChart->Name);

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Charts'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Chart count: ' . $loadedSheet->Charts->Count . PHP_EOL;
        echo 'First chart: ' . $loadedSheet->Charts[0]->Name . ' (Type: ' . self::ChartTypeName($loadedSheet->Charts[0]->ChartType) . ')' . PHP_EOL;
        echo 'Second chart: ' . $loadedSheet->Charts[1]->Name . ' (Type: ' . self::ChartTypeName($loadedSheet->Charts[1]->ChartType) . ')' . PHP_EOL;
        echo 'First chart anchor: R' . $loadedSheet->Charts[0]->UpperLeftRow . 'C' . $loadedSheet->Charts[0]->UpperLeftColumn . ' to R' . $loadedSheet->Charts[0]->LowerRightRow . 'C' . $loadedSheet->Charts[0]->LowerRightColumn . PHP_EOL;
    }

    private static function ChartTypeName($value)
    {
        switch ($value) {
            case ChartType::Bar:
                return 'Bar';
            case ChartType::Column:
                return 'Column';
            case ChartType::Line:
                return 'Line';
            case ChartType::Area:
                return 'Area';
            case ChartType::Pie:
                return 'Pie';
            case ChartType::Doughnut:
                return 'Doughnut';
            case ChartType::Scatter:
                return 'Scatter';
            case ChartType::Bubble:
                return 'Bubble';
            case ChartType::Radar:
                return 'Radar';
            case ChartType::Stock:
                return 'Stock';
            case ChartType::Bar3D:
                return 'Bar3D';
            case ChartType::Column3D:
                return 'Column3D';
            case ChartType::Line3D:
                return 'Line3D';
            case ChartType::Area3D:
                return 'Area3D';
            case ChartType::Pie3D:
                return 'Pie3D';
            case ChartType::Surface3D:
                return 'Surface3D';
            case ChartType::SurfaceWireframe3D:
                return 'SurfaceWireframe3D';
            case ChartType::Contour:
                return 'Contour';
            case ChartType::Waterfall:
                return 'Waterfall';
            case ChartType::Treemap:
                return 'Treemap';
            case ChartType::Sunburst:
                return 'Sunburst';
            case ChartType::Histogram:
                return 'Histogram';
            case ChartType::BoxAndWhisker:
                return 'BoxAndWhisker';
            case ChartType::Funnel:
                return 'Funnel';
            case ChartType::Map:
                return 'Map';
            case ChartType::Unknown:
                return 'Unknown';
            default:
                return (string)$value;
        }
    }

    private static function ResolveOutputDirectory()
    {
        $directory = self::ResolveRepositoryRoot() . DIRECTORY_SEPARATOR . 'output' . DIRECTORY_SEPARATOR . 'samples';
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        return $directory;
    }

    private static function ResolveRepositoryRoot()
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== '') {
            if (is_file($directory . DIRECTORY_SEPARATOR . 'composer.json')) {
                return $directory;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }

            $directory = $parent;
        }

        return getcwd();
    }
}

Program::Main();
