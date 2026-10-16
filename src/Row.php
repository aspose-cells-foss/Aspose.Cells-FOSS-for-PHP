<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\RowModel;

final class Row
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
        if ($propertyName === 'Height') {
            $model = $this->TryGetModel();
            return $model === null ? null : $model->Height;
        }

        if ($propertyName === 'IsHidden') {
            $model = $this->TryGetModel();
            return $model === null ? false : $model->Hidden;
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Height') {
            if ($value !== null && $value <= 0.0) {
                throw new CellsException("Row height must be positive.");
            }

            $model = $this->GetOrCreateModel();
            $model->Height = $value;
            $this->Normalize();
            return;
        }

        if ($propertyName === 'IsHidden') {
            $model = $this->GetOrCreateModel();
            $model->Hidden = $value;
            $this->Normalize();
            return;
        }
    }

    private function TryGetModel()
    {
        $model = null;
        if (array_key_exists($this->_index, $this->_worksheet->Model->Rows)) {
            $model = $this->_worksheet->Model->Rows[$this->_index];
        }
        return $model;
    }

    private function GetOrCreateModel()
    {
        if (array_key_exists($this->_index, $this->_worksheet->Model->Rows)) {
            return $this->_worksheet->Model->Rows[$this->_index];
        }

        $created = new RowModel();
        $this->_worksheet->Model->Rows[$this->_index] = $created;
        return $created;
    }

    private function Normalize()
    {
        if (array_key_exists($this->_index, $this->_worksheet->Model->Rows)) {
            $model = $this->_worksheet->Model->Rows[$this->_index];
            if ($model->Height === null && !$model->Hidden && $model->StyleIndex === null) {
                unset($this->_worksheet->Model->Rows[$this->_index]);
            }
        }
    }
}
