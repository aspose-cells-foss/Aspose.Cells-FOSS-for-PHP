<?php

namespace Aspose\Cells_FOSS;

final class FilterValueCollection implements \ArrayAccess, \Countable, \IteratorAggregate
{
    private $values;

    public function __construct(&$values)
    {
        $this->values = &$values;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->values);
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
        return $offset >= 0 && $offset < count($this->values);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0 || $offset >= count($this->values)) {
            throw new CellsException('Filter value index was out of range.');
        }

        return $this->values[$offset];
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
        return count($this->values);
    }

    public function Add($value)
    {
        $normalized = AutoFilterSupport::NormalizeText($value, 'value');
        $this->values[] = $normalized;
        return count($this->values) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->values)) {
            throw new CellsException('Filter value index was out of range.');
        }

        array_splice($this->values, $index, 1);
    }

    public function Clear()
    {
        $this->values = array();
    }

    public function GetEnumerator()
    {
        $values = array();
        for ($index = 0; $index < count($this->values); $index++) {
            $values[] = $this->values[$index];
        }

        return $values;
    }

    #[\ReturnTypeWillChange]
    public function getIterator()
    {
        return new \ArrayIterator($this->GetEnumerator());
    }
}
