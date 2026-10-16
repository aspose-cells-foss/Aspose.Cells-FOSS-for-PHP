<?php

namespace Aspose\Cells_FOSS;

final class LoadIssue
{
    public $Code;
    public $Severity;
    public $Message;
    public $RepairApplied;
    public $DataLossRisk;
    public $PartUri;
    public $SheetName;
    public $CellRef;
    public $RowIndex;

    public function __construct($code, $severity, $message, $repairApplied = false, $dataLossRisk = false)
    {
        if ($code === null) throw new \InvalidArgumentException('code');
        $this->Code = $code;
        $this->Severity = $severity;
        if ($message === null) throw new \InvalidArgumentException('message');
        $this->Message = $message;
        $this->RepairApplied = $repairApplied;
        $this->DataLossRisk = $dataLossRisk;
    }
}