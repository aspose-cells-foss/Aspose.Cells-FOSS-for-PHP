<?php

namespace Aspose\Cells_FOSS\Rendering;

interface PdfTextFontBackend
{
    /**
     * @param PdfTextRunRecord[] $runs
     */
    public function Initialize($runs);

    /**
     * @param int &$nextObjectId
     * @return array<int, string> objectId => object bytes
     */
    public function BuildObjects(&$nextObjectId);

    /**
     * @return string PDF font resource dictionary fragment (without enclosing << >>).
     */
    public function BuildFontResourceText();

    /**
     * @param PdfTextRunRecord $run
     * @return PdfEncodedTextSegment[]
     */
    public function EncodeRun($run);

    public function Dispose();
}
