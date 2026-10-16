<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\ConditionalFormattingModel;
use Aspose\Cells_FOSS\Core\FormatConditionModel;
use Aspose\Cells_FOSS\Core\StyleValue;

// Represents a collection of format condition objects.
final class FormatConditionCollection implements \ArrayAccess, \Countable
{
    private $_owner;
    private $_model;

    public function __construct(&$owner, ConditionalFormattingModel $model)
    {
        $this->_owner =& $owner;
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_model->Conditions);
        }
        if ($propertyName === 'RangeCount') {
            return count($this->_model->Areas);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Count') {
            throw new \RuntimeException('Count is read-only.');
        }
        if ($propertyName === 'RangeCount') {
            throw new \RuntimeException('RangeCount is read-only.');
        }
    }

    public function count(): int
    {
        return count($this->_model->Conditions);
    }

    public function offsetExists($offset): bool
    {
        $index = (int) $offset;
        return $index >= 0 && $index < count($this->_model->Conditions);
    }

    public function offsetGet($offset): mixed
    {
        $index = (int) $offset;
        if ($index < 0 || $index >= count($this->_model->Conditions)) {
            throw new CellsException('Format condition index was out of range.');
        }

        return new FormatCondition($this->_owner, $this->_model, $this->_model->Conditions[$index]);
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The format condition collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The format condition collection is read-only.');
    }

    public function Add($area, $type, $operatorType, $formula1, $formula2)
    {
        $this->AddArea($area);
        return $this->AddCondition($type, $operatorType, $formula1, $formula2);
    }

    public function AddCondition($type, $operatorType = OperatorType::None, $formula1 = '', $formula2 = '')
    {
        $condition = new FormatConditionModel();
        $condition->Type = $type;
        $condition->Operator = $operatorType;
        $condition->Formula1 = self::NormalizeFormula($formula1);
        $condition->Formula2 = self::NormalizeFormula($formula2);
        $condition->Priority = ConditionalFormattingCollection::GetNextPriority($this->_owner);
        $condition->Style = StyleValue::Default()->Clone();
        self::InitializeDefaults($condition);
        $this->_model->Conditions[] = $condition;
        return count($this->_model->Conditions) - 1;
    }

    public function AddArea($area)
    {
        self::ValidateArea($area);
        $this->_model->Areas[] = $area;
        self::SortAreas($this->_model->Areas);
    }

    public function GetCellArea($index)
    {
        if ($index < 0 || $index >= count($this->_model->Areas)) {
            throw new CellsException('Conditional formatting area index was out of range.');
        }

        return $this->_model->Areas[$index];
    }

    public function RemoveArea($a, $b = null, $c = null, $d = null)
    {
        if (func_num_args() === 1) {
            if (is_object($a)) {
                self::ValidateArea($a);
                self::ReplaceAreas($this->_model, self::SubtractAreas($this->_model->Areas, $a));
                self::RemoveCollectionIfEmpty($this->_owner, $this->_model);
                return;
            }

            $index = $a;
            if ($index < 0 || $index >= count($this->_model->Areas)) {
                throw new CellsException('Conditional formatting area index was out of range.');
            }

            array_splice($this->_model->Areas, $index, 1);
            self::RemoveCollectionIfEmpty($this->_owner, $this->_model);
            return;
        }

        self::RemoveArea(CellArea::CreateCellArea($a, $b, $c, $d));
    }

    public function RemoveCondition($ownerOrIndex, $collection = null, $model = null)
    {
        if ($collection !== null) {
            $owner = $ownerOrIndex;
            for ($i = 0; $i < count($collection->Conditions); $i++) {
                if ($collection->Conditions[$i] === $model) {
                    array_splice($collection->Conditions, $i, 1);
                    break;
                }
            }

            self::RemoveCollectionIfEmpty($owner, $collection);
            return;
        }

        $index = $ownerOrIndex;
        if ($index < 0 || $index >= count($this->_model->Conditions)) {
            throw new CellsException('Format condition index was out of range.');
        }

        array_splice($this->_model->Conditions, $index, 1);
        self::RemoveCollectionIfEmpty($this->_owner, $this->_model);
    }

    private static function InitializeDefaults($condition)
    {
        switch ($condition->Type) {
            case FormatConditionType::DuplicateValues:
                $condition->Duplicate = true;
                break;
            case FormatConditionType::UniqueValues:
                $condition->Duplicate = false;
                break;
            case FormatConditionType::Top10:
                $condition->Top = true;
                $condition->Rank = 10;
                break;
            case FormatConditionType::Bottom10:
                $condition->Top = false;
                $condition->Rank = 10;
                break;
            case FormatConditionType::AboveAverage:
                $condition->Above = true;
                break;
            case FormatConditionType::BelowAverage:
                $condition->Above = false;
                break;
            case FormatConditionType::ColorScale:
                $condition->ColorScaleCount = 2;
                break;
            case FormatConditionType::DataBar:
                $condition->BarColor = new ColorValue(255, 99, 142, 198);
                break;
            case FormatConditionType::IconSet:
                $condition->IconSetType = '3TrafficLights1';
                break;
        }
    }

    private static function RemoveCollectionIfEmpty($owner, $collection)
    {
        if (count($collection->Areas) === 0 || count($collection->Conditions) === 0) {
            for ($i = 0; $i < count($owner); $i++) {
                if ($owner[$i] === $collection) {
                    array_splice($owner, $i, 1);
                    break;
                }
            }
        }
    }

    private static function ReplaceAreas($model, $areas)
    {
        $model->Areas = array();
        for ($index = 0; $index < count($areas); $index++) {
            $model->Areas[] = $areas[$index];
        }
    }

    private static function NormalizeFormula($value)
    {
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if ($trimmed === '') {
            return null;
        }

        if ($trimmed[0] === '=') {
            return substr($trimmed, 1);
        }

        return $trimmed;
    }

    private static function ValidateArea($area)
    {
        if ($area->FirstRow < 0 || $area->FirstColumn < 0 || $area->TotalRows <= 0 || $area->TotalColumns <= 0) {
            throw new CellsException('Conditional formatting area must be a positive cell range.');
        }
    }

    private static function SubtractAreas($sourceAreas, $removal)
    {
        $remaining = array();
        for ($index = 0; $index < count($sourceAreas); $index++) {
            self::SubtractArea($sourceAreas[$index], $removal, $remaining);
        }

        self::SortAreas($remaining);
        return $remaining;
    }

    private static function SubtractArea($source, $removal, &$output)
    {
        if (!self::AreasOverlap($source, $removal)) {
            $output[] = $source;
            return;
        }

        $sourceLastRow = $source->FirstRow + $source->TotalRows - 1;
        $sourceLastColumn = $source->FirstColumn + $source->TotalColumns - 1;
        $removalLastRow = $removal->FirstRow + $removal->TotalRows - 1;
        $removalLastColumn = $removal->FirstColumn + $removal->TotalColumns - 1;

        $overlapFirstRow = max($source->FirstRow, $removal->FirstRow);
        $overlapFirstColumn = max($source->FirstColumn, $removal->FirstColumn);
        $overlapLastRow = min($sourceLastRow, $removalLastRow);
        $overlapLastColumn = min($sourceLastColumn, $removalLastColumn);

        self::AddIfNonEmpty($output, $source->FirstRow, $source->FirstColumn, $overlapFirstRow - 1, $sourceLastColumn);
        self::AddIfNonEmpty($output, $overlapLastRow + 1, $source->FirstColumn, $sourceLastRow, $sourceLastColumn);
        self::AddIfNonEmpty($output, $overlapFirstRow, $source->FirstColumn, $overlapLastRow, $overlapFirstColumn - 1);
        self::AddIfNonEmpty($output, $overlapFirstRow, $overlapLastColumn + 1, $overlapLastRow, $sourceLastColumn);
    }

    private static function AddIfNonEmpty(&$areas, $firstRow, $firstColumn, $lastRow, $lastColumn)
    {
        if ($lastRow < $firstRow || $lastColumn < $firstColumn) {
            return;
        }

        $areas[] = CellArea::CreateCellArea($firstRow, $firstColumn, $lastRow, $lastColumn);
    }

    public static function AreasOverlap($left, $right)
    {
        $leftLastRow = $left->FirstRow + $left->TotalRows - 1;
        $leftLastColumn = $left->FirstColumn + $left->TotalColumns - 1;
        $rightLastRow = $right->FirstRow + $right->TotalRows - 1;
        $rightLastColumn = $right->FirstColumn + $right->TotalColumns - 1;

        return $left->FirstRow <= $rightLastRow
            && $right->FirstRow <= $leftLastRow
            && $left->FirstColumn <= $rightLastColumn
            && $right->FirstColumn <= $leftLastColumn;
    }

    public static function SortAreas(&$areas)
    {
        usort($areas, array(__CLASS__, 'CompareAreas'));
    }

    public static function CompareAreas($left, $right)
    {
        $rowComparison = $left->FirstRow <=> $right->FirstRow;
        if ($rowComparison != 0) {
            return $rowComparison;
        }

        $columnComparison = $left->FirstColumn <=> $right->FirstColumn;
        if ($columnComparison != 0) {
            return $columnComparison;
        }

        $rowCountComparison = $left->TotalRows <=> $right->TotalRows;
        if ($rowCountComparison != 0) {
            return $rowCountComparison;
        }

        return $left->TotalColumns <=> $right->TotalColumns;
    }
}