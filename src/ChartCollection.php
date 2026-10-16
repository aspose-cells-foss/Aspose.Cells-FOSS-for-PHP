<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ChartModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;

// Represents collection of charts on a worksheet.
final class ChartCollection implements \ArrayAccess, \Countable
{
    private $_model;

    public function __construct(WorksheetModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_model->Charts);
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
        return count($this->_model->Charts);
    }

    public function offsetExists($offset): bool
    {
        $index = (int) $offset;
        return $index >= 0 && $index < count($this->_model->Charts);
    }

    public function offsetGet($offset): mixed
    {
        $index = (int) $offset;
        if ($index < 0 || $index >= count($this->_model->Charts)) {
            throw new CellsException('Chart index ' . $index . ' is out of range. The worksheet contains ' . count($this->_model->Charts) . ' chart(s).');
        }

        return new Chart($this->_model->Charts[$index]);
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The chart collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The chart collection is read-only.');
    }

    public function Add($type, $dataRange, $upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn)
    {
        $chartId = count($this->_model->Charts) + 1;
        $rawXml = ChartXmlTemplates::Build($type, $dataRange, 'Chart ' . $chartId);
        $model = new ChartModel();
        $model->Name = 'Chart ' . $chartId;
        $model->ChartType = $type;
        $model->UpperLeftRow = $upperLeftRow;
        $model->UpperLeftColumn = $upperLeftColumn;
        $model->LowerRightRow = $lowerRightRow;
        $model->LowerRightColumn = $lowerRightColumn;
        $model->RawChartXml = $rawXml;
        $this->_model->Charts[] = $model;
        return count($this->_model->Charts) - 1;
    }
}