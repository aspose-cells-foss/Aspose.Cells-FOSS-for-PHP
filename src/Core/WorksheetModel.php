<?php

namespace Aspose\Cells_FOSS\Core;

final class WorksheetModel
{
    public $Name;
    public $Visibility = SheetVisibility::Visible;
    public $Cells;
    public $Rows;
    public $Columns;
    public $MergeRegions;
    public $Hyperlinks;
    public $Validations;
    public $ConditionalFormattings;
    public $PageSetup;
    public $View;
    public $Protection;
    public $AutoFilter;
    public $ListObjects;
    public $Pictures;
    public $ShapeImages;
    public $Shapes;
    public $Charts;
    public $Comments;
    public $SmartArts;
    public $TabColor = null;
    public $BaseColumnWidth = null;
    public $DefaultColumnWidth = null;
    public $DefaultRowHeight = null;
    public $CustomHeight = null;

    public function __construct($name)
    {
        $this->Name = $name;
        $this->Cells = array();
        $this->Rows = array();
        $this->Columns = array();
        $this->MergeRegions = array();
        $this->Hyperlinks = array();
        $this->Validations = array();
        $this->ConditionalFormattings = array();
        $this->PageSetup = new PageSetupModel();
        $this->View = new WorksheetViewModel();
        $this->Protection = new WorksheetProtectionModel();
        $this->AutoFilter = new AutoFilterModel();
        $this->ListObjects = array();
        $this->Pictures = array();
        $this->ShapeImages = array();
        $this->Shapes = array();
        $this->Charts = array();
        $this->Comments = array();
        $this->SmartArts = array();
    }
}