<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKRect;

final class PdfTextDocumentSession
{
    private $runs = array();

    public function HasRuns()
    {
        return count($this->runs) > 0;
    }

    /**
     * @return PdfTextRunRecord[]
     */
    public function &Runs()
    {
        return $this->runs;
    }

    public function RecordWorksheetRun($page, $clipRect, $text, $typeface, $color, $fontSizePt, $xPt, $baselinePt)
    {
        $this->RecordPageRun($page, $clipRect, $text, $typeface, $color, $fontSizePt, $xPt, $baselinePt);
    }

    public function RecordChartRun($page, $clipRect, $text, $typeface, $color, $fontSizePt, $xPt, $baselinePt, $invisible = false)
    {
        if ($page === null || $typeface === null || $text === null || $text === '') {
            return;
        }

        $clipRect = PageContentClip::Intersect($page, $clipRect, true);
        if ($clipRect->IsEmpty) {
            return;
        }

        $scale = (float)$page->ScaleFactor;
        $x = (float)$page->ContentOriginXPt + $scale * $xPt;
        $baselineTop = (float)$page->ContentOriginYPt + $scale * $baselinePt;
        $clipLeft = (float)$page->ContentOriginXPt + $scale * $clipRect->Left;
        $clipTop = (float)$page->ContentOriginYPt + $scale * $clipRect->Top;
        $clipRight = (float)$page->ContentOriginXPt + $scale * $clipRect->Right;
        $clipBottom = (float)$page->ContentOriginYPt + $scale * $clipRect->Bottom;

        $record = new PdfTextRunRecord();
        $record->PageNumber = $page->PageNumber;
        $record->Text = $text;
        $record->Typeface = $typeface;
        $record->UsePositiveYFontMatrix = true;
        $record->UseTopDownCoordinates = true;
        $record->Color = $color;
        $record->Invisible = $invisible;
        $record->FontSizePt = $scale * $fontSizePt;
        $record->XPt = $x;
        $record->BaselineYPt = $baselineTop;
        $record->PageHeightPt = (float)$page->PageHeightPt;
        $record->ClipLeftPt = $clipLeft;
        $record->ClipTopPt = $clipTop;
        $record->ClipBottomPt = $clipBottom;
        $record->ClipWidthPt = $clipRight - $clipLeft;
        $record->ClipHeightPt = $clipBottom - $clipTop;
        $this->runs[] = $record;
    }

    public function RecordChartRotatedRun($page, $clipRect, $text, $typeface, $color, $fontSizePt, $originXPt, $originYPt, $localXPt, $localBaselinePt, $rotationDeg)
    {
        if ($page === null || $typeface === null || $text === null || $text === '') {
            return;
        }

        $clipRect = PageContentClip::Intersect($page, $clipRect, true);
        if ($clipRect->IsEmpty) {
            return;
        }

        $scale = (float)$page->ScaleFactor;
        $originX = (float)$page->ContentOriginXPt + $scale * $originXPt;
        $originY = (float)$page->ContentOriginYPt + $scale * $originYPt;
        $clipLeft = (float)$page->ContentOriginXPt + $scale * $clipRect->Left;
        $clipTop = (float)$page->ContentOriginYPt + $scale * $clipRect->Top;
        $clipRight = (float)$page->ContentOriginXPt + $scale * $clipRect->Right;
        $clipBottom = (float)$page->ContentOriginYPt + $scale * $clipRect->Bottom;

        $record = new PdfTextRunRecord();
        $record->PageNumber = $page->PageNumber;
        $record->Text = $text;
        $record->Typeface = $typeface;
        $record->UsePositiveYFontMatrix = true;
        $record->UseTopDownCoordinates = true;
        $record->Color = $color;
        $record->FontSizePt = $scale * $fontSizePt;
        $record->XPt = $scale * $localXPt;
        $record->BaselineYPt = $scale * $localBaselinePt;
        $record->PageHeightPt = (float)$page->PageHeightPt;
        $record->ClipLeftPt = $clipLeft;
        $record->ClipTopPt = $clipTop;
        $record->ClipBottomPt = $clipBottom;
        $record->ClipWidthPt = $clipRight - $clipLeft;
        $record->ClipHeightPt = $clipBottom - $clipTop;
        $record->TransformOriginXPt = $originX;
        $record->TransformOriginYPt = $originY;
        $record->RotationDeg = $rotationDeg;
        $this->runs[] = $record;
    }

    private function RecordPageRun($page, $clipRect, $text, $typeface, $color, $fontSizePt, $xPt, $baselinePt)
    {
        if ($page === null || $typeface === null || $text === null || $text === '') {
            return;
        }

        $clipRect = PageContentClip::Intersect($page, $clipRect, false);
        if ($clipRect->IsEmpty) {
            return;
        }

        $scale = (float)$page->ScaleFactor;
        $x = (float)$page->ContentOriginXPt + $scale * $xPt;
        $pageHeight = (float)$page->PageHeightPt;
        $baselineTopDown = (float)$page->ContentOriginYPt + $scale * $baselinePt;
        $baseline = $pageHeight - $baselineTopDown;

        $clipLeft = (float)$page->ContentOriginXPt + $scale * $clipRect->Left;
        $clipTop = (float)$page->ContentOriginYPt + $scale * $clipRect->Top;
        $clipRight = (float)$page->ContentOriginXPt + $scale * $clipRect->Right;
        $clipBottom = (float)$page->ContentOriginYPt + $scale * $clipRect->Bottom;
        $clipBottomPdf = $pageHeight - $clipBottom;

        $record = new PdfTextRunRecord();
        $record->PageNumber = $page->PageNumber;
        $record->Text = $text;
        $record->Typeface = $typeface;
        $record->Color = $color;
        $record->FontSizePt = $scale * $fontSizePt;
        $record->XPt = $x;
        $record->BaselineYPt = $baseline;
        $record->PageHeightPt = $pageHeight;
        $record->ClipLeftPt = $clipLeft;
        $record->ClipTopPt = $clipTop;
        $record->ClipBottomPt = $clipBottomPdf;
        $record->ClipWidthPt = $clipRight - $clipLeft;
        $record->ClipHeightPt = $clipBottom - $clipTop;
        $this->runs[] = $record;
    }
}
