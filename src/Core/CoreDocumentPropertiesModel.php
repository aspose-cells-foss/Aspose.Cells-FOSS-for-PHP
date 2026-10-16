<?php

namespace Aspose\Cells_FOSS\Core;

final class CoreDocumentPropertiesModel
{
    public $Title = '';
    public $Subject = '';
    public $Creator = '';
    public $Keywords = '';
    public $Description = '';
    public $LastModifiedBy = '';
    public $Revision = '';
    public $Category = '';
    public $ContentStatus = '';
    public $Created = null;
    public $Modified = null;

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->Title = $source->Title;
        $this->Subject = $source->Subject;
        $this->Creator = $source->Creator;
        $this->Keywords = $source->Keywords;
        $this->Description = $source->Description;
        $this->LastModifiedBy = $source->LastModifiedBy;
        $this->Revision = $source->Revision;
        $this->Category = $source->Category;
        $this->ContentStatus = $source->ContentStatus;
        $this->Created = $source->Created;
        $this->Modified = $source->Modified;
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return !($this->Title === null || $this->Title === '')
            || !($this->Subject === null || $this->Subject === '')
            || !($this->Creator === null || $this->Creator === '')
            || !($this->Keywords === null || $this->Keywords === '')
            || !($this->Description === null || $this->Description === '')
            || !($this->LastModifiedBy === null || $this->LastModifiedBy === '')
            || !($this->Revision === null || $this->Revision === '')
            || !($this->Category === null || $this->Category === '')
            || !($this->ContentStatus === null || $this->ContentStatus === '')
            || $this->Created !== null
            || $this->Modified !== null;
    }
}
