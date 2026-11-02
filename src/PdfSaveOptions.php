<?php

namespace Aspose\Cells_FOSS;

final class PdfSaveOptions extends SaveOptions
{
    public $OnePagePerSheet = false;

    public $AllColumnsInOnePagePerSheet = false;

    public $DefaultFont = null;

    public $ConditionalFormattingEvaluationDate = null;

    public $UseType3TextOptimization;

    public $UseTrueTypeSubsetPrototype;

    public function __construct()
    {
        $this->SaveFormat = SaveFormat::Pdf;
        $this->UseType3TextOptimization = true;
        $this->UseTrueTypeSubsetPrototype = true;
    }
}
