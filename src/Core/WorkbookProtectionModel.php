<?php

namespace Aspose\Cells_FOSS\Core;

final class WorkbookProtectionModel
{
    public $LockStructure = false;
    public $LockWindows = false;
    public $LockRevision = false;
    public $WorkbookPassword = '';
    public $RevisionsPassword = '';

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->LockStructure = $source->LockStructure;
        $this->LockWindows = $source->LockWindows;
        $this->LockRevision = $source->LockRevision;
        $this->WorkbookPassword = $source->WorkbookPassword;
        $this->RevisionsPassword = $source->RevisionsPassword;
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return $this->LockStructure
            || $this->LockWindows
            || $this->LockRevision
            || !($this->WorkbookPassword === null || $this->WorkbookPassword === '')
            || !($this->RevisionsPassword === null || $this->RevisionsPassword === '');
    }
}