<?php

namespace Aspose\Cells_FOSS\Core;

final class ExtendedDocumentPropertiesModel
{
    public $Application = '';
    public $AppVersion = '';
    public $Company = '';
    public $Manager = '';
    public $DocSecurity = null;
    public $HyperlinkBase = '';
    public $ScaleCrop = null;
    public $LinksUpToDate = null;
    public $SharedDoc = null;

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->Application = $source->Application;
        $this->AppVersion = $source->AppVersion;
        $this->Company = $source->Company;
        $this->Manager = $source->Manager;
        $this->DocSecurity = $source->DocSecurity;
        $this->HyperlinkBase = $source->HyperlinkBase;
        $this->ScaleCrop = $source->ScaleCrop;
        $this->LinksUpToDate = $source->LinksUpToDate;
        $this->SharedDoc = $source->SharedDoc;
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return !($this->Application === null || $this->Application === '')
            || !($this->AppVersion === null || $this->AppVersion === '')
            || !($this->Company === null || $this->Company === '')
            || !($this->Manager === null || $this->Manager === '')
            || $this->DocSecurity !== null
            || !($this->HyperlinkBase === null || $this->HyperlinkBase === '')
            || $this->ScaleCrop !== null
            || $this->LinksUpToDate !== null
            || $this->SharedDoc !== null;
    }
}
