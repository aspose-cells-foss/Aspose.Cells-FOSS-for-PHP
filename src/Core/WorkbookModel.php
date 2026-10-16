<?php

namespace Aspose\Cells_FOSS\Core;

final class WorkbookModel
{
    public $Worksheets;
    public $Settings;
    public $Properties;
    public $DocumentProperties;
    public $Diagnostics;
    public $Styles;
    public $SharedStrings;
    public $DefaultStyle;
    public $ActiveSheetIndex;
    public $DefinedNames;
    public $RawThemeXml;
    public $ExternalLinks;

    public function __construct()
    {
        $this->Settings = new WorkbookSettingsModel();
        $this->Properties = new WorkbookPropertiesModel();
        $this->DocumentProperties = new DocumentPropertiesModel();
        $this->Diagnostics = new DiagnosticBag();
        $this->Styles = new StyleRepository();
        $this->SharedStrings = new SharedStringRepository();
        $this->DefaultStyle = StyleValue::Default()->Clone();
        $this->Worksheets = array(
            new WorksheetModel('Sheet1'),
        );
        $this->DefinedNames = array();
        $this->ActiveSheetIndex = 0;
        $this->ExternalLinks = array();
    }
}