<?php

namespace Aspose\Cells_FOSS;

final class RowCollection implements \ArrayAccess
{
    private $_worksheet;

    public function __construct($worksheet)
    {
        $this->_worksheet = $worksheet;
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($index)
    {
        if ($index < 0) throw new CellsException("Row index must be non-negative.");
        return new Row($this->_worksheet, $index);
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($index, $value)
    {
        throw new \RuntimeException('RowCollection indexer is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($index)
    {
        throw new \RuntimeException('RowCollection indexer is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($index)
    {
        try {
            $this->offsetGet($index);
            return true;
        } catch (\Throwable $exception) {
            return false;
        }
    }
}
