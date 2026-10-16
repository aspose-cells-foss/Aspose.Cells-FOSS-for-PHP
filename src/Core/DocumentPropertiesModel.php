<?php

namespace Aspose\Cells_FOSS\Core;

final class DocumentPropertiesModel
{
    public $Core;
    public $Extended;

    public function __construct()
    {
        $this->Core = new CoreDocumentPropertiesModel();
        $this->Extended = new ExtendedDocumentPropertiesModel();
    }

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->Core->CopyFrom($source->Core);
        $this->Extended->CopyFrom($source->Extended);
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return $this->Core->HasStoredState() || $this->Extended->HasStoredState();
    }
}
