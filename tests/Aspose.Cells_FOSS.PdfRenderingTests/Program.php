<?php

namespace Aspose\Cells_FOSS\PdfRenderingTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';

use Aspose\Cells_FOSS\ChartType;
use Aspose\Cells_FOSS\PdfSaveOptions;
use Aspose\Cells_FOSS\SaveFormat;
use Aspose\Cells_FOSS\Workbook;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;

final class Program
{
    public static function Main()
    {
        return TestRunner::Run(
            'PdfRendering.ConversionTests',
            array(
                new TestCase('pdf_save_options_defaults_to_pdf_format', array(Program::class, 'PdfSaveOptionsDefaultsToPdfFormat')),
                new TestCase('workbook_saves_valid_pdf_document', array(Program::class, 'WorkbookSavesValidPdfDocument')),
                new TestCase('save_format_pdf_enum_produces_pdf', array(Program::class, 'SaveFormatPdfEnumProducesPdf')),
                new TestCase('empty_workbook_emits_blank_page', array(Program::class, 'EmptyWorkbookEmitsBlankPage')),
                new TestCase('chart_workbook_saves_valid_pdf_document', array(Program::class, 'ChartWorkbookSavesValidPdfDocument')),
            )
        );
    }

    public static function PdfSaveOptionsDefaultsToPdfFormat()
    {
        $options = new PdfSaveOptions();
        AssertEx::Equal(SaveFormat::Pdf, $options->SaveFormat);
    }

    public static function WorkbookSavesValidPdfDocument()
    {
        $directory = new TemporaryDirectory('pdf-rendering');
        $path = $directory->GetPath('values.pdf');

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells['A1']->PutValue('Item');
        $sheet->Cells['B1']->PutValue('Price');
        $sheet->Cells['A2']->PutValue('Apple');
        $sheet->Cells['B2']->PutValue(2.5);
        $sheet->Cells['A3']->PutValue('Banana');
        $sheet->Cells['B3']->PutValue(1.75);
        $workbook->Save($path);

        self::AssertPdf($path);
        $directory->Dispose();
    }

    public static function SaveFormatPdfEnumProducesPdf()
    {
        $directory = new TemporaryDirectory('pdf-rendering');
        $path = $directory->GetPath('save-format.pdf');

        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells['A1']->PutValue('Hello PDF');
        $workbook->Save($path, SaveFormat::Pdf);

        self::AssertPdf($path);
        $directory->Dispose();
    }

    public static function EmptyWorkbookEmitsBlankPage()
    {
        $directory = new TemporaryDirectory('pdf-rendering');
        $path = $directory->GetPath('empty.pdf');

        $workbook = new Workbook();
        $workbook->Save($path, new PdfSaveOptions());

        self::AssertPdf($path);
        $directory->Dispose();
    }

    public static function ChartWorkbookSavesValidPdfDocument()
    {
        $directory = new TemporaryDirectory('pdf-rendering');
        $path = $directory->GetPath('chart.pdf');

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Cells['A1']->PutValue('Month');
        $sheet->Cells['B1']->PutValue('Sales');
        for ($month = 1; $month <= 6; $month++) {
            $sheet->Cells[[$month, 0]]->PutValue('Month ' . $month);
            $sheet->Cells[[$month, 1]]->PutValue($month * 100);
        }

        $sheet->Charts->Add(ChartType::Column, 'Sheet1!$B$1:$B$7', 0, 4, 16, 12);
        $workbook->Save($path, new PdfSaveOptions());

        self::AssertPdf($path);
        $directory->Dispose();
    }

    private static function AssertPdf($path)
    {
        AssertEx::True(is_file($path), 'Expected the PDF file to be created: ' . $path);
        $bytes = file_get_contents($path);
        AssertEx::True(str_starts_with($bytes, '%PDF-'), 'Expected a PDF header.');
        AssertEx::True(str_contains($bytes, '%%EOF'), 'Expected a PDF trailer.');
        AssertEx::True(strlen($bytes) > 0, 'Expected non-empty PDF content.');
    }
}

Program::Main();
