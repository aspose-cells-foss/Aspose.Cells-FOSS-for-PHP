<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlReadContext
{
    public function __construct($archive, $sharedStrings, $dateSystem, $styles, $sheets)
    {
        $this->Archive = $archive;
        $this->SharedStrings = $sharedStrings;
        $this->DateSystem = $dateSystem;
        $this->Styles = $styles;
        $this->Sheets = $sheets;
    }

    public $Archive;
    public $SharedStrings;
    public $DateSystem;
    public $Styles;
    public $Sheets;
}
