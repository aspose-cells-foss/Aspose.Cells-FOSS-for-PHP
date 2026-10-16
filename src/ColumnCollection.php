<?php

namespace Aspose\Cells_FOSS;

final class ColumnCollection implements \ArrayAccess
{
    private $_worksheet;

    public function __construct($worksheet)
    {
        $this->_worksheet = $worksheet;
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0) throw new CellsException("Column index must be non-negative.");
        return new Column($this->_worksheet, $offset);
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        throw new \RuntimeException('ColumnCollection indexer is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        throw new \RuntimeException('ColumnCollection indexer is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        try {
            $this->offsetGet($offset);
            return true;
        } catch (\Throwable $exception) {
            return false;
        }
    }
}
