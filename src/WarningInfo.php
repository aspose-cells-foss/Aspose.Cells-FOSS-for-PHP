<?php

namespace Aspose\Cells_FOSS;

final class WarningInfo
{
    public $Code = '';
    public $Severity = DiagnosticSeverity::Warning;
    public $Message = '';
    public $DataLossRisk = false;
    public $PartUri;
    public $SheetName;
    public $CellRef;
    public $RowIndex;
}