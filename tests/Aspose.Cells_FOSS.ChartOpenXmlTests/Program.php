<?php

namespace Aspose\Cells_FOSS\ChartOpenXmlTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';
require_once __DIR__ . '/../Shared/ChartScenarioFactory.php';

use Aspose\Cells_FOSS\ChartType;
use Aspose\Cells_FOSS\SaveFormat;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\ChartScenarioFactory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        return TestRunner::Run(
            'Chart.OpenXmlTests',
            array(
                new TestCase('column_chart_saves_bar_element_with_col_direction', array(Program::class, 'ColumnChartSavesBarElementWithColDirection')),
                new TestCase('bar_chart_saves_bar_element_with_bar_direction', array(Program::class, 'BarChartSavesBarElementWithBarDirection')),
                new TestCase('line_chart_saves_line_chart_element', array(Program::class, 'LineChartSavesLineChartElement')),
                new TestCase('area_chart_saves_area_chart_element', array(Program::class, 'AreaChartSavesAreaChartElement')),
                new TestCase('pie_chart_saves_pie_chart_element', array(Program::class, 'PieChartSavesPieChartElement')),
                new TestCase('doughnut_chart_saves_doughnut_chart_element', array(Program::class, 'DoughnutChartSavesDoughnutChartElement')),
                new TestCase('scatter_chart_saves_scatter_chart_element', array(Program::class, 'ScatterChartSavesScatterChartElement')),
                new TestCase('radar_chart_saves_radar_chart_element', array(Program::class, 'RadarChartSavesRadarChartElement')),
                new TestCase('bubble_chart_saves_bubble_chart_element', array(Program::class, 'BubbleChartSavesBubbleChartElement')),
                new TestCase('stock_chart_saves_stock_chart_element', array(Program::class, 'StockChartSavesStockChartElement')),
                new TestCase('column3d_chart_saves_bar3d_element_with_col', array(Program::class, 'Column3DChartSavesBar3DElementWithCol')),
                new TestCase('bar3d_chart_saves_bar3d_element_with_bar', array(Program::class, 'Bar3DChartSavesBar3DElementWithBar')),
                new TestCase('line3d_chart_saves_line3d_element', array(Program::class, 'Line3DChartSavesLine3DElement')),
                new TestCase('area3d_chart_saves_area3d_element', array(Program::class, 'Area3DChartSavesArea3DElement')),
                new TestCase('pie3d_chart_saves_pie3d_element', array(Program::class, 'Pie3DChartSavesPie3DElement')),
                new TestCase('surface3d_chart_saves_surface3d_element', array(Program::class, 'Surface3DChartSavesSurface3DElement')),
                new TestCase('contour_chart_saves_surface_element', array(Program::class, 'ContourChartSavesSurfaceElement')),
                new TestCase('chart_drawing_xml_contains_graphic_frame', array(Program::class, 'ChartDrawingXmlContainsGraphicFrame')),
                new TestCase('chart_content_types_includes_chart_override', array(Program::class, 'ChartContentTypesIncludesChartOverride')),
                new TestCase('chart_drawing_rels_links_to_chart_file', array(Program::class, 'ChartDrawingRelsLinksToChartFile')),
                new TestCase('multiple_charts_on_same_sheet_all_saved', array(Program::class, 'MultipleChartsOnSameSheetAllSaved')),
                new TestCase('chart_preserves_data_range_in_xml', array(Program::class, 'ChartPreservesDataRangeInXml')),
            )
        );
    }

    // ---- per-type XML structure tests ----

    public static function ColumnChartSavesBarElementWithColDirection()
    {
        $chartXml = self::GetChartXml(ChartType::Column);
        AssertEx::Contains('<c:barChart', $chartXml, 'Expected c:barChart element');
        AssertEx::Contains('barDir val="col"', $chartXml, 'Expected barDir=col for column chart');
    }

    public static function BarChartSavesBarElementWithBarDirection()
    {
        $chartXml = self::GetChartXml(ChartType::Bar);
        AssertEx::Contains('<c:barChart', $chartXml, 'Expected c:barChart element');
        AssertEx::Contains('barDir val="bar"', $chartXml, 'Expected barDir=bar for bar chart');
    }

    public static function LineChartSavesLineChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Line);
        AssertEx::Contains('<c:lineChart', $chartXml, 'Expected c:lineChart element');
    }

    public static function AreaChartSavesAreaChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Area);
        AssertEx::Contains('<c:areaChart', $chartXml, 'Expected c:areaChart element');
    }

    public static function PieChartSavesPieChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Pie);
        AssertEx::Contains('<c:pieChart', $chartXml, 'Expected c:pieChart element');
    }

    public static function DoughnutChartSavesDoughnutChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Doughnut);
        AssertEx::Contains('<c:doughnutChart', $chartXml, 'Expected c:doughnutChart element');
    }

    public static function ScatterChartSavesScatterChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Scatter);
        AssertEx::Contains('<c:scatterChart', $chartXml, 'Expected c:scatterChart element');
    }

    public static function RadarChartSavesRadarChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Radar);
        AssertEx::Contains('<c:radarChart', $chartXml, 'Expected c:radarChart element');
    }

    public static function BubbleChartSavesBubbleChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Bubble);
        AssertEx::Contains('<c:bubbleChart', $chartXml, 'Expected c:bubbleChart element');
    }

    public static function StockChartSavesStockChartElement()
    {
        $chartXml = self::GetChartXml(ChartType::Stock);
        AssertEx::Contains('<c:stockChart', $chartXml, 'Expected c:stockChart element');
    }

    public static function Column3DChartSavesBar3DElementWithCol()
    {
        $chartXml = self::GetChartXml(ChartType::Column3D);
        AssertEx::Contains('<c:bar3DChart', $chartXml, 'Expected c:bar3DChart element');
        AssertEx::Contains('barDir val="col"', $chartXml, 'Expected barDir=col for column3D chart');
    }

    public static function Bar3DChartSavesBar3DElementWithBar()
    {
        $chartXml = self::GetChartXml(ChartType::Bar3D);
        AssertEx::Contains('<c:bar3DChart', $chartXml, 'Expected c:bar3DChart element');
        AssertEx::Contains('barDir val="bar"', $chartXml, 'Expected barDir=bar for bar3D chart');
    }

    public static function Line3DChartSavesLine3DElement()
    {
        $chartXml = self::GetChartXml(ChartType::Line3D);
        AssertEx::Contains('<c:line3DChart', $chartXml, 'Expected c:line3DChart element');
    }

    public static function Area3DChartSavesArea3DElement()
    {
        $chartXml = self::GetChartXml(ChartType::Area3D);
        AssertEx::Contains('<c:area3DChart', $chartXml, 'Expected c:area3DChart element');
    }

    public static function Pie3DChartSavesPie3DElement()
    {
        $chartXml = self::GetChartXml(ChartType::Pie3D);
        AssertEx::Contains('<c:pie3DChart', $chartXml, 'Expected c:pie3DChart element');
    }

    public static function Surface3DChartSavesSurface3DElement()
    {
        $chartXml = self::GetChartXml(ChartType::Surface3D);
        AssertEx::Contains('<c:surface3DChart', $chartXml, 'Expected c:surface3DChart element');
    }

    public static function ContourChartSavesSurfaceElement()
    {
        $chartXml = self::GetChartXml(ChartType::Contour);
        AssertEx::Contains('<c:surfaceChart', $chartXml, 'Expected c:surfaceChart element');
    }

    public static function ChartDrawingXmlContainsGraphicFrame()
    {
        $stream = self::SaveChartWorkbook(ChartType::Column);
        $drawingXml = self::ReadZipEntry($stream, 'xl/drawings/drawing1.xml');
        AssertEx::Contains('graphicFrame', $drawingXml, 'Expected graphicFrame in drawing XML');
        AssertEx::Contains('http://schemas.openxmlformats.org/drawingml/2006/chart', $drawingXml,
            'Drawing XML must reference chart namespace URI');
    }

    public static function ChartContentTypesIncludesChartOverride()
    {
        $stream = self::SaveChartWorkbook(ChartType::Line);
        $contentTypes = self::ReadZipEntry($stream, '[Content_Types].xml');
        AssertEx::Contains('/xl/charts/chart1.xml', $contentTypes,
            'Content_Types.xml must have chart1.xml override');
        AssertEx::Contains('application/vnd.openxmlformats-officedocument.drawingml.chart+xml', $contentTypes,
            'Content_Types.xml must have chart content type');
    }

    public static function ChartDrawingRelsLinksToChartFile()
    {
        $stream = self::SaveChartWorkbook(ChartType::Pie);
        $relsXml = self::ReadZipEntry($stream, 'xl/drawings/_rels/drawing1.xml.rels');
        AssertEx::Contains('../charts/chart1.xml', $relsXml,
            'Drawing rels must link to chart file');
        AssertEx::Contains('http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', $relsXml,
            'Drawing rels must have chart relationship type');
    }

    public static function MultipleChartsOnSameSheetAllSaved()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Charts->Add(ChartType::Column, 'Sheet1!$B$1:$B$5', 0, 0, 10, 8);
        $sheet->Charts->Add(ChartType::Line, 'Sheet1!$C$1:$C$5', 12, 0, 22, 8);

        $stream = fopen('php://temp', 'r+b');
        $workbook->Save($stream, SaveFormat::Xlsx);
        rewind($stream);
        $reloaded = new Workbook($stream);
        AssertEx::Equal(2, $reloaded->Worksheets[0]->Charts->Count, 'Expected two charts after reload');
        AssertEx::Equal(ChartType::Column, $reloaded->Worksheets[0]->Charts[0]->ChartType, 'First chart must be column');
        AssertEx::Equal(ChartType::Line, $reloaded->Worksheets[0]->Charts[1]->ChartType, 'Second chart must be line');

        $contentTypes = self::ReadZipEntry($stream, '[Content_Types].xml');
        AssertEx::Contains('/xl/charts/chart1.xml', $contentTypes, 'Content_Types.xml must have chart1.xml override');
        AssertEx::Contains('/xl/charts/chart2.xml', $contentTypes, 'Content_Types.xml must have chart2.xml override');
    }

    public static function ChartPreservesDataRangeInXml()
    {
        $dataRange = 'Sheet1!$B$1:$B$5';
        $workbook = ChartScenarioFactory::CreateChartWorkbook(ChartType::Column);

        $stream = fopen('php://temp', 'r+b');
        $workbook->Save($stream, SaveFormat::Xlsx);
        $chartXml = self::ReadZipEntry($stream, 'xl/charts/chart1.xml');
        AssertEx::Contains($dataRange, $chartXml, 'Chart XML must preserve the data range');
    }

    // ---- helpers ----

    private static function GetChartXml($type)
    {
        $stream = self::SaveChartWorkbook($type);
        return self::ReadZipEntry($stream, 'xl/charts/chart1.xml');
    }

    private static function SaveChartWorkbook($type)
    {
        $workbook = ChartScenarioFactory::CreateChartWorkbook($type);

        // Write to disk for manual inspection
        $outputDir = self::GetOrCreateOutputDirectory('chart-openxml-roundtrips');
        $workbook->Save($outputDir . DIRECTORY_SEPARATOR . self::GetChartTypeName($type) . '.xlsx');

        $stream = fopen('php://temp', 'r+b');
        $workbook->Save($stream, SaveFormat::Xlsx);
        return $stream;
    }

    private static function GetOrCreateOutputDirectory($subdir)
    {
        $directory = __DIR__;
        while ($directory !== dirname($directory)) {
            if (file_exists($directory . DIRECTORY_SEPARATOR . 'composer.json')) {
                $outputDir = $directory . DIRECTORY_SEPARATOR . 'output' . DIRECTORY_SEPARATOR . $subdir;
                if (!is_dir($outputDir)) {
                    mkdir($outputDir, 0777, true);
                }
                return $outputDir;
            }
            $directory = dirname($directory);
        }

        throw new \RuntimeException('Could not locate repository root.');
    }

    private static function ReadZipEntry($stream, $entryName)
    {
        rewind($stream);
        $content = stream_get_contents($stream);

        $tempFile = tempnam(sys_get_temp_dir(), 'chartzip');
        file_put_contents($tempFile, $content);

        $archive = new \ZipArchive();
        if ($archive->open($tempFile) !== true) {
            unlink($tempFile);
            throw new \RuntimeException('Could not open zip package.');
        }

        $normalized = ltrim($entryName, '/');
        $text = $archive->getFromName($normalized);
        $archive->close();
        unlink($tempFile);

        if ($text === false) {
            throw new \RuntimeException('Missing zip entry: ' . $entryName);
        }

        return $text;
    }

    private static function GetChartTypeName($type)
    {
        $names = array(
            ChartType::Bar => 'Bar',
            ChartType::Column => 'Column',
            ChartType::Line => 'Line',
            ChartType::Area => 'Area',
            ChartType::Pie => 'Pie',
            ChartType::Doughnut => 'Doughnut',
            ChartType::Scatter => 'Scatter',
            ChartType::Bubble => 'Bubble',
            ChartType::Radar => 'Radar',
            ChartType::Stock => 'Stock',
            ChartType::Bar3D => 'Bar3D',
            ChartType::Column3D => 'Column3D',
            ChartType::Line3D => 'Line3D',
            ChartType::Area3D => 'Area3D',
            ChartType::Pie3D => 'Pie3D',
            ChartType::Surface3D => 'Surface3D',
            ChartType::SurfaceWireframe3D => 'SurfaceWireframe3D',
            ChartType::Contour => 'Contour',
            ChartType::Waterfall => 'Waterfall',
            ChartType::Treemap => 'Treemap',
            ChartType::Sunburst => 'Sunburst',
            ChartType::Histogram => 'Histogram',
            ChartType::BoxAndWhisker => 'BoxAndWhisker',
            ChartType::Funnel => 'Funnel',
            ChartType::Map => 'Map',
            ChartType::Unknown => 'Unknown',
        );

        return isset($names[$type]) ? $names[$type] : (string)$type;
    }
}

Program::Main();
