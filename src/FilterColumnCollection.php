<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\FilterColumnModel;

final class FilterColumnCollection implements \ArrayAccess, \Countable, \IteratorAggregate
{
    private $models;

    public function __construct(&$models)
    {
        $this->models = &$models;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->models);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Count') {
            throw new \RuntimeException('Count is read-only.');
        }
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        return $offset >= 0 && $offset < count($this->models);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0 || $offset >= count($this->models)) {
            throw new CellsException('Filter column index was out of range.');
        }

        return new FilterColumn($this->models[$offset]);
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        throw new \RuntimeException('Collection is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        throw new \RuntimeException('Collection is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function count()
    {
        return count($this->models);
    }

    public function Add($columnIndex)
    {
        if ($columnIndex < 0) {
            throw new CellsException('Filter column index must be zero or greater.');
        }

        for ($index = 0; $index < count($this->models); $index++) {
            if ($this->models[$index]->ColumnIndex === $columnIndex) {
                throw new CellsException('A filter column for the specified column index already exists.');
            }
        }

        $model = new FilterColumnModel();
        $model->ColumnIndex = $columnIndex;

        $insertIndex = 0;
        while ($insertIndex < count($this->models) && $this->models[$insertIndex]->ColumnIndex < $columnIndex) {
            $insertIndex++;
        }

        array_splice($this->models, $insertIndex, 0, array($model));
        return $insertIndex;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->models)) {
            throw new CellsException('Filter column index was out of range.');
        }

        array_splice($this->models, $index, 1);
    }

    public function Clear()
    {
        $this->models = array();
    }

    public function GetEnumerator()
    {
        $columns = array();
        for ($index = 0; $index < count($this->models); $index++) {
            $columns[] = new FilterColumn($this->models[$index]);
        }

        return $columns;
    }

    #[\ReturnTypeWillChange]
    public function getIterator()
    {
        return new \ArrayIterator($this->GetEnumerator());
    }
}
