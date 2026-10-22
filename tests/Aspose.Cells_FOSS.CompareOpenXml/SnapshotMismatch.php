<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class SnapshotMismatch
{
    public function __construct($cellKey, $librarySnapshot, $openXmlSnapshot)
    {
        $this->CellKey = $cellKey;
        $this->LibrarySnapshot = $librarySnapshot;
        $this->OpenXmlSnapshot = $openXmlSnapshot;
    }

    public $CellKey;
    public $LibrarySnapshot;
    public $OpenXmlSnapshot;
}
