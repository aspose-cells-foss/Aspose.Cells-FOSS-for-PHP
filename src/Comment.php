<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CommentModel;

// Represents a worksheet comment (legacy note) anchored to a single cell.
final class Comment
{
    private $_model;

    public function __construct(CommentModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Row') {
            return $this->_model->Row;
        }
        if ($propertyName === 'Column') {
            return $this->_model->Column;
        }
        if ($propertyName === 'Author') {
            return $this->_model->Author;
        }
        if ($propertyName === 'Note') {
            return $this->_model->Note;
        }
        if ($propertyName === 'IsVisible') {
            return $this->_model->IsVisible;
        }
        if ($propertyName === 'Width') {
            return $this->_model->Width;
        }
        if ($propertyName === 'Height') {
            return $this->_model->Height;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Row') {
            throw new \RuntimeException('Row is read-only.');
        }
        if ($propertyName === 'Column') {
            throw new \RuntimeException('Column is read-only.');
        }
        if ($propertyName === 'Author') {
            $this->_model->Author = $value === null ? '' : $value;
            return;
        }
        if ($propertyName === 'Note') {
            $this->_model->Note = $value === null ? '' : $value;
            return;
        }
        if ($propertyName === 'IsVisible') {
            $this->_model->IsVisible = $value;
            return;
        }
        if ($propertyName === 'Width') {
            if ($value < 1) {
                throw new CellsException('Comment Width must be at least 1 pixel.');
            }

            $this->_model->Width = $value;
            return;
        }
        if ($propertyName === 'Height') {
            if ($value < 1) {
                throw new CellsException('Comment Height must be at least 1 pixel.');
            }

            $this->_model->Height = $value;
            return;
        }
    }
}