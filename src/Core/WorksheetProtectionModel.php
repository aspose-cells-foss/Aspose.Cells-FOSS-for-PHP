<?php

namespace Aspose\Cells_FOSS\Core;

final class WorksheetProtectionModel
{
    public $IsProtected = false;
    public $Objects = false;
    public $Scenarios = false;
    public $FormatCells = false;
    public $FormatColumns = false;
    public $FormatRows = false;
    public $InsertColumns = false;
    public $InsertRows = false;
    public $InsertHyperlinks = false;
    public $DeleteColumns = false;
    public $DeleteRows = false;
    public $SelectLockedCells = false;
    public $Sort = false;
    public $AutoFilter = false;
    public $PivotTables = false;
    public $SelectUnlockedCells = false;
    public $PasswordHash;
    public $AlgorithmName;
    public $HashValue;
    public $SaltValue;
    public $SpinCount;

    // Clears the current state.
    public function Clear()
    {
        $this->IsProtected = false;
        $this->Objects = false;
        $this->Scenarios = false;
        $this->FormatCells = false;
        $this->FormatColumns = false;
        $this->FormatRows = false;
        $this->InsertColumns = false;
        $this->InsertRows = false;
        $this->InsertHyperlinks = false;
        $this->DeleteColumns = false;
        $this->DeleteRows = false;
        $this->SelectLockedCells = false;
        $this->Sort = false;
        $this->AutoFilter = false;
        $this->PivotTables = false;
        $this->SelectUnlockedCells = false;
        $this->PasswordHash = null;
        $this->AlgorithmName = null;
        $this->HashValue = null;
        $this->SaltValue = null;
        $this->SpinCount = null;
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return $this->IsProtected
            || $this->Objects
            || $this->Scenarios
            || $this->FormatCells
            || $this->FormatColumns
            || $this->FormatRows
            || $this->InsertColumns
            || $this->InsertRows
            || $this->InsertHyperlinks
            || $this->DeleteColumns
            || $this->DeleteRows
            || $this->SelectLockedCells
            || $this->Sort
            || $this->AutoFilter
            || $this->PivotTables
            || $this->SelectUnlockedCells
            || !($this->PasswordHash === null || trim($this->PasswordHash) === '')
            || !($this->AlgorithmName === null || trim($this->AlgorithmName) === '')
            || !($this->HashValue === null || trim($this->HashValue) === '')
            || !($this->SaltValue === null || trim($this->SaltValue) === '')
            || !($this->SpinCount === null || trim($this->SpinCount) === '');
    }
}