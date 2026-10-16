<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ListObjectModel;

// Represents the ordered collection of columns in an Excel table.
final class ListColumnCollection implements \ArrayAccess, \Countable
{
    private $_model;

    public function __construct(ListObjectModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_model->Columns);
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
        return count($this->_model->Columns);
    }

    public function offsetExists($offset): bool
    {
        $index = (int) $offset;
        return $index >= 0 && $index < count($this->_model->Columns);
    }

    public function offsetGet($offset): mixed
    {
        $index = (int) $offset;
        if ($index < 0 || $index >= count($this->_model->Columns)) {
            throw new CellsException('Column index ' . $index . ' is out of range.');
        }

        return new ListColumn($this->_model->Columns[$index]);
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The list column collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The list column collection is read-only.');
    }
}