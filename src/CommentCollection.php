<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CommentModel;

// Represents the collection of comments (legacy notes) on a worksheet.
final class CommentCollection implements \ArrayAccess, \Countable
{
    private $_models;

    public function __construct(&$models)
    {
        $this->_models =& $models;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_models);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Count') {
            throw new \RuntimeException('Count is read-only.');
        }
    }

    public function count(): int
    {
        return count($this->_models);
    }

    public function offsetExists($offset): bool
    {
        if (is_int($offset)) {
            return $offset >= 0 && $offset < count($this->_models);
        }

        $address = null;
        if (!self::TryParseCell((string) $offset, $address)) {
            return false;
        }

        for ($i = 0; $i < count($this->_models); $i++) {
            if ($this->_models[$i]->Row === $address->RowIndex && $this->_models[$i]->Column === $address->ColumnIndex) {
                return true;
            }
        }

        return false;
    }

    public function offsetGet($offset): mixed
    {
        if (is_int($offset)) {
            if ($offset < 0 || $offset >= count($this->_models)) {
                throw new CellsException('Comment index ' . $offset . ' is out of range.');
            }

            return new Comment($this->_models[$offset]);
        }

        $address = null;
        if (!self::TryParseCell((string) $offset, $address)) {
            return null;
        }

        for ($i = 0; $i < count($this->_models); $i++) {
            if ($this->_models[$i]->Row === $address->RowIndex && $this->_models[$i]->Column === $address->ColumnIndex) {
                return new Comment($this->_models[$i]);
            }
        }

        return null;
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The comment collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The comment collection is read-only.');
    }

    public function Add($rowOrCellName, $column = null)
    {
        if ($column === null) {
            $cellName = $rowOrCellName;
            $address = null;
            if (!self::TryParseCell($cellName, $address)) {
                throw new CellsException("'" . $cellName . "' is not a valid A1-style cell reference.");
            }

            return $this->Add($address->RowIndex, $address->ColumnIndex);
        }

        $row = $rowOrCellName;
        if ($row < 0) {
            throw new CellsException('Comment row must be non-negative.');
        }

        if ($column < 0) {
            throw new CellsException('Comment column must be non-negative.');
        }

        for ($i = 0; $i < count($this->_models); $i++) {
            if ($this->_models[$i]->Row === $row && $this->_models[$i]->Column === $column) {
                throw new CellsException('A comment already exists at row ' . $row . ', column ' . $column . '.');
            }
        }

        $model = new CommentModel();
        $model->Row = $row;
        $model->Column = $column;
        $this->_models[] = $model;
        return new Comment($model);
    }

    public function RemoveAt($indexOrCellName)
    {
        if (is_int($indexOrCellName)) {
            $index = $indexOrCellName;
            if ($index < 0 || $index >= count($this->_models)) {
                throw new CellsException('Comment index ' . $index . ' is out of range.');
            }

            array_splice($this->_models, $index, 1);
            return;
        }

        $cellName = $indexOrCellName;
        $address = null;
        if (!self::TryParseCell($cellName, $address)) {
            return;
        }

        for ($i = 0; $i < count($this->_models); $i++) {
            if ($this->_models[$i]->Row === $address->RowIndex && $this->_models[$i]->Column === $address->ColumnIndex) {
                array_splice($this->_models, $i, 1);
                return;
            }
        }
    }

    private static function TryParseCell($cellName, &$address)
    {
        try {
            $address = CellAddress::Parse($cellName);
            return true;
        } catch (\InvalidArgumentException $exception) {
            $address = null;
            return false;
        }
    }
}