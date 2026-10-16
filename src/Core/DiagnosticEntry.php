<?php
namespace Aspose\Cells_FOSS\Core;

final class DiagnosticEntry
{
    public $Code;
    public $Severity;
    public $Message;
    public $RepairApplied;
    public $DataLossRisk;

    public function __construct()
    {
        $this->Code = "";
        $this->Severity = DiagnosticSeverity::Warning;
        $this->Message = "";
    }
}
