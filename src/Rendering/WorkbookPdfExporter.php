<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\SheetVisibility;
use Aspose\Cells_FOSS\PdfSaveOptions;
use Aspose\Cells_FOSS\Skia\SKDocument;
use Aspose\Cells_FOSS\Skia\SKManagedWStream;

final class WorkbookPdfExporter
{
    private function __construct()
    {
    }

    public static function Export($workbook, $stream, $options = null)
    {
        if ($workbook === null) {
            throw new \InvalidArgumentException('workbook');
        }

        if ($stream === null) {
            throw new \InvalidArgumentException('stream');
        }

        if ($options === null) {
            $options = new PdfSaveOptions();
        }

        $context = new RenderContext($workbook, $options->DefaultFont);

        try {
            if ($options->ConditionalFormattingEvaluationDate !== null) {
                $date = $options->ConditionalFormattingEvaluationDate;
                if ($date instanceof \DateTime) {
                    $date = $date->format('Y-m-d');
                } elseif ($date instanceof \DateTimeInterface) {
                    $date = $date->format('Y-m-d');
                }

                $context->ConditionalFormattingEvaluationDate = new \DateTime($date);
            }

            if ($options->UseType3TextOptimization) {
                // The PDF text post-processing chain is not ported; text is always drawn as
                // vector outlines through PdfTextPathRenderer.
                $context->PdfTextSession = null;
                $context->EnableWorksheetTextOptimization = false;
            }

            $pages = self::BuildPages($context, $workbook, $options);

            for ($i = 0; $i < count($pages); $i++) {
                $pages[$i]->PageNumber = $i + 1;
                $pages[$i]->TotalPages = count($pages);
            }

            $pdfBuffer = new SKManagedWStream();

            try {
                $document = SKDocument::CreatePdf($pdfBuffer);
                $renderer = new SheetPdfRenderer($context);

                if (count($pages) === 0) {
                    // A valid PDF needs at least one page; emit a blank default (A4) page.
                    PaperSizes::GetDimensionsPoints(0, $w, $h);
                    $document->BeginPage((float)$w, (float)$h);
                    $document->EndPage();
                }

                foreach ($pages as $page) {
                    $canvas = $document->BeginPage((float)$page->PageWidthPt, (float)$page->PageHeightPt);
                    $renderer->RenderPage($canvas, $page);
                    $document->EndPage();
                }

                $document->Close();

                $bytes = $pdfBuffer->ToBytes();
            } finally {
                $pdfBuffer->Dispose();
            }

            self::WriteAll($stream, $bytes);
        } finally {
            $context->Dispose();
        }
    }

    private static function BuildPages($context, $workbook, $options)
    {
        $pages = array();

        for ($sheetIndex = 0; $sheetIndex < count($workbook->Worksheets); $sheetIndex++) {
            $sheet = $workbook->Worksheets[$sheetIndex];
            if ($sheet->Visibility !== SheetVisibility::Visible) {
                continue;
            }

            $layout = SheetLayout::Build($context, $sheet);
            if ($layout->IsEmpty) {
                continue;
            }

            foreach (Paginator::Paginate($layout, $sheetIndex, $options) as $page) {
                $pages[] = $page;
            }
        }

        return $pages;
    }

    private static function WriteAll($stream, $bytes)
    {
        if (is_string($stream)) {
            file_put_contents($stream, $bytes);
            return;
        }

        if (is_resource($stream)) {
            fwrite($stream, $bytes);
            return;
        }

        if (is_object($stream) && method_exists($stream, 'Write')) {
            $stream->Write($bytes);
            return;
        }

        throw new \InvalidArgumentException('stream');
    }
}
