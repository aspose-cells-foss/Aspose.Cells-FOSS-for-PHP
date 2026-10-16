<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ShapeModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;

// Represents collection of drawing objects (shapes) on a worksheet.
final class ShapeCollection implements \ArrayAccess, \Countable
{
    private $_worksheetModel;

    public function __construct(WorksheetModel $worksheetModel)
    {
        $this->_worksheetModel = $worksheetModel;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_worksheetModel->Shapes);
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
        return count($this->_worksheetModel->Shapes);
    }

    public function offsetExists($offset): bool
    {
        $index = (int) $offset;
        return $index >= 0 && $index < count($this->_worksheetModel->Shapes);
    }

    public function offsetGet($offset): mixed
    {
        $index = (int) $offset;
        if ($index < 0 || $index >= count($this->_worksheetModel->Shapes)) {
            throw new CellsException('Shape index ' . $index . ' is out of range.');
        }

        return new Shape($this->_worksheetModel->Shapes[$index]);
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The shape collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The shape collection is read-only.');
    }

    public function Add($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn, $shapeType)
    {
        if ($upperLeftRow < 0) {
            throw new CellsException('upperLeftRow must be non-negative.');
        }

        if ($upperLeftColumn < 0) {
            throw new CellsException('upperLeftColumn must be non-negative.');
        }

        if ($lowerRightRow < $upperLeftRow) {
            throw new CellsException('lowerRightRow must be >= upperLeftRow.');
        }

        if ($lowerRightColumn < $upperLeftColumn) {
            throw new CellsException('lowerRightColumn must be >= upperLeftColumn.');
        }

        $index = count($this->_worksheetModel->Shapes);
        $model = new ShapeModel();
        $model->Name = 'Shape ' . ($index + 1);
        $model->UpperLeftRow = $upperLeftRow;
        $model->UpperLeftColumn = $upperLeftColumn;
        $model->LowerRightRow = $lowerRightRow;
        $model->LowerRightColumn = $lowerRightColumn;
        $model->GeometryType = Shape::AutoShapeTypeToGeometry($shapeType);
        $this->_worksheetModel->Shapes[] = $model;
        return $index;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->_worksheetModel->Shapes)) {
            throw new CellsException('Shape index ' . $index . ' is out of range.');
        }

        array_splice($this->_worksheetModel->Shapes, $index, 1);
    }
}