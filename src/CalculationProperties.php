<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CalculationPropertiesModel;

final class CalculationProperties
{
    private $_model;

    public function __construct(CalculationPropertiesModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'CalculationId') {
            return $this->_model->CalculationId;
        }
        if ($propertyName === 'CalculationMode') {
            if ($this->_model->CalculationMode === null || $this->_model->CalculationMode === '') {
                return 'auto';
            }
            return $this->_model->CalculationMode;
        }
        if ($propertyName === 'FullCalculationOnLoad') {
            return $this->_model->FullCalculationOnLoad;
        }
        if ($propertyName === 'ReferenceMode') {
            if ($this->_model->ReferenceMode === null || $this->_model->ReferenceMode === '') {
                return 'A1';
            }
            return $this->_model->ReferenceMode;
        }
        if ($propertyName === 'Iterate') {
            return $this->_model->Iterate;
        }
        if ($propertyName === 'IterateCount') {
            return $this->_model->IterateCount ?? 100;
        }
        if ($propertyName === 'IterateDelta') {
            return $this->_model->IterateDelta ?? 0.001;
        }
        if ($propertyName === 'FullPrecision') {
            return $this->_model->FullPrecision === null || $this->_model->FullPrecision;
        }
        if ($propertyName === 'CalculationCompleted') {
            return $this->_model->CalculationCompleted === null || $this->_model->CalculationCompleted;
        }
        if ($propertyName === 'CalculationOnSave') {
            return $this->_model->CalculationOnSave === null || $this->_model->CalculationOnSave;
        }
        if ($propertyName === 'ConcurrentCalculation') {
            return $this->_model->ConcurrentCalculation === null || $this->_model->ConcurrentCalculation;
        }
        if ($propertyName === 'ForceFullCalculation') {
            return $this->_model->ForceFullCalculation;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'CalculationId') {
            if ($value !== null && $value < 0) {
                throw new CellsException('CalculationId must be non-negative.');
            }
            $this->_model->CalculationId = $value;
            return;
        }
        if ($propertyName === 'CalculationMode') {
            $this->_model->CalculationMode = WorkbookPropertySupport::NormalizeCalculationMode($value);
            return;
        }
        if ($propertyName === 'FullCalculationOnLoad') {
            $this->_model->FullCalculationOnLoad = $value;
            return;
        }
        if ($propertyName === 'ReferenceMode') {
            $this->_model->ReferenceMode = WorkbookPropertySupport::NormalizeReferenceMode($value);
            return;
        }
        if ($propertyName === 'Iterate') {
            $this->_model->Iterate = $value;
            return;
        }
        if ($propertyName === 'IterateCount') {
            if ($value < 0) {
                throw new CellsException('IterateCount must be non-negative.');
            }
            $this->_model->IterateCount = $value;
            return;
        }
        if ($propertyName === 'IterateDelta') {
            if ($value < 0.0) {
                throw new CellsException('IterateDelta must be non-negative.');
            }
            $this->_model->IterateDelta = $value;
            return;
        }
        if ($propertyName === 'FullPrecision') {
            $this->_model->FullPrecision = $value;
            return;
        }
        if ($propertyName === 'CalculationCompleted') {
            $this->_model->CalculationCompleted = $value;
            return;
        }
        if ($propertyName === 'CalculationOnSave') {
            $this->_model->CalculationOnSave = $value;
            return;
        }
        if ($propertyName === 'ConcurrentCalculation') {
            $this->_model->ConcurrentCalculation = $value;
            return;
        }
        if ($propertyName === 'ForceFullCalculation') {
            $this->_model->ForceFullCalculation = $value;
            return;
        }
    }
}
