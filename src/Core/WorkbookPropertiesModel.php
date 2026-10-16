<?php

namespace Aspose\Cells_FOSS\Core;

final class WorkbookPropertiesModel
{
    public $CodeName = '';
    public $ShowObjects = '';
    public $FilterPrivacy = false;
    public $ShowBorderUnselectedTables = true;
    public $ShowInkAnnotation = true;
    public $BackupFile = false;
    public $SaveExternalLinkValues = true;
    public $UpdateLinks = '';
    public $HidePivotFieldList = false;
    public $DefaultThemeVersion = null;
    public $Protection;
    public $View;
    public $Calculation;

    public function __construct()
    {
        $this->Protection = new WorkbookProtectionModel();
        $this->View = new WorkbookViewModel();
        $this->Calculation = new CalculationPropertiesModel();
    }

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->CodeName = $source->CodeName;
        $this->ShowObjects = $source->ShowObjects;
        $this->FilterPrivacy = $source->FilterPrivacy;
        $this->ShowBorderUnselectedTables = $source->ShowBorderUnselectedTables;
        $this->ShowInkAnnotation = $source->ShowInkAnnotation;
        $this->BackupFile = $source->BackupFile;
        $this->SaveExternalLinkValues = $source->SaveExternalLinkValues;
        $this->UpdateLinks = $source->UpdateLinks;
        $this->HidePivotFieldList = $source->HidePivotFieldList;
        $this->DefaultThemeVersion = $source->DefaultThemeVersion;
        $this->Protection->CopyFrom($source->Protection);
        $this->View->CopyFrom($source->View);
        $this->Calculation->CopyFrom($source->Calculation);
    }

    // Performs has workbook properties state.
    public function HasWorkbookPropertiesState()
    {
        return !($this->CodeName === null || $this->CodeName === '')
            || !($this->ShowObjects === null || $this->ShowObjects === '')
            || $this->FilterPrivacy
            || !$this->ShowBorderUnselectedTables
            || !$this->ShowInkAnnotation
            || $this->BackupFile
            || !$this->SaveExternalLinkValues
            || !($this->UpdateLinks === null || $this->UpdateLinks === '')
            || $this->HidePivotFieldList
            || $this->DefaultThemeVersion !== null;
    }
}