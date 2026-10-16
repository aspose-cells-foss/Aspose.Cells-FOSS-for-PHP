<?php

namespace Aspose\Cells_FOSS\Core;

final class CalculationPropertiesModel
{
    public $CalculationId = null;
    public $CalculationMode = '';
    public $FullCalculationOnLoad = false;
    public $ReferenceMode = '';
    public $Iterate = false;
    public $IterateCount = null;
    public $IterateDelta = null;
    public $FullPrecision = null;
    public $CalculationCompleted = null;
    public $CalculationOnSave = null;
    public $ConcurrentCalculation = null;
    public $ForceFullCalculation = false;

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->CalculationId = $source->CalculationId;
        $this->CalculationMode = $source->CalculationMode;
        $this->FullCalculationOnLoad = $source->FullCalculationOnLoad;
        $this->ReferenceMode = $source->ReferenceMode;
        $this->Iterate = $source->Iterate;
        $this->IterateCount = $source->IterateCount;
        $this->IterateDelta = $source->IterateDelta;
        $this->FullPrecision = $source->FullPrecision;
        $this->CalculationCompleted = $source->CalculationCompleted;
        $this->CalculationOnSave = $source->CalculationOnSave;
        $this->ConcurrentCalculation = $source->ConcurrentCalculation;
        $this->ForceFullCalculation = $source->ForceFullCalculation;
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return $this->CalculationId !== null
            || !($this->CalculationMode === null || $this->CalculationMode === '')
            || $this->FullCalculationOnLoad
            || !($this->ReferenceMode === null || $this->ReferenceMode === '')
            || $this->Iterate
            || $this->IterateCount !== null
            || $this->IterateDelta !== null
            || $this->FullPrecision !== null
            || $this->CalculationCompleted !== null
            || $this->CalculationOnSave !== null
            || $this->ConcurrentCalculation !== null
            || $this->ForceFullCalculation;
    }
}
