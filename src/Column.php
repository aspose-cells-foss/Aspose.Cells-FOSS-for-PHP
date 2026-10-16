<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ColumnRangeModel;

final class Column
{
    private $_worksheet;
    private $_index;

    public function __construct($worksheet, $index)
    {
        $this->_worksheet = $worksheet;
        $this->_index = $index;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Width') {
            $model = $this->FindModel();
            return $model === null ? null : $model->Width;
        }

        if ($propertyName === 'IsHidden') {
            $model = $this->FindModel();
            return $model === null ? false : $model->Hidden;
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Width') {
            if ($value !== null && $value <= 0.0) {
                throw new CellsException("Column width must be positive.");
            }

            $model = $this->GetOrCreateModel();
            $model->Width = $value;
            $this->Normalize($model);
            return;
        }

        if ($propertyName === 'IsHidden') {
            $model = $this->GetOrCreateModel();
            $model->Hidden = $value;
            $this->Normalize($model);
            return;
        }
    }

    private function FindModel()
    {
        for ($index = count($this->_worksheet->Model->Columns) - 1; $index >= 0; $index--) {
            $model = $this->_worksheet->Model->Columns[$index];
            if ($model->MinColumnIndex <= $this->_index && $model->MaxColumnIndex >= $this->_index) {
                return $model;
            }
        }

        return null;
    }

    private function GetOrCreateModel()
    {
        $existing = $this->FindModel();
        if ($existing !== null && $existing->MinColumnIndex == $this->_index && $existing->MaxColumnIndex == $this->_index) {
            return $existing;
        }

        $created = new ColumnRangeModel();
        $created->MinColumnIndex = $this->_index;
        $created->MaxColumnIndex = $this->_index;
        $this->_worksheet->Model->Columns[] = $created;
        return $created;
    }

    private function Normalize($model)
    {
        if ($model->Width === null && !$model->Hidden && $model->StyleIndex === null) {
            $index = array_search($model, $this->_worksheet->Model->Columns, true);
            if ($index !== false) {
                array_splice($this->_worksheet->Model->Columns, $index, 1);
            }
        }
    }
}
