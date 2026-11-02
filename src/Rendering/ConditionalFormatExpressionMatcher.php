<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\CellArea;
use Aspose\Cells_FOSS\Core\CellAddress;

final class ConditionalFormatExpressionMatcher
{
    private const Comparison = '/\A\s*(\$?[A-Z]{1,3}\$?[1-9][0-9]*)\s*(>=|<=|<>|>|<|=)\s*([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?)\s*\z/i';

    private $_sheet;
    private $_anchorRow;
    private $_anchorColumn;
    private $_referenceRow;
    private $_referenceColumn;
    private $_absoluteRow;
    private $_absoluteColumn;
    private $_operator;
    private $_threshold;

    private function __construct($sheet, $area, $reference, $absoluteRow, $absoluteColumn,
        $comparisonOperator, $threshold)
    {
        $this->_sheet = $sheet;
        $this->_anchorRow = $area->StartRow;
        $this->_anchorColumn = $area->StartColumn;
        $this->_referenceRow = $reference->RowIndex;
        $this->_referenceColumn = $reference->ColumnIndex;
        $this->_absoluteRow = $absoluteRow;
        $this->_absoluteColumn = $absoluteColumn;
        $this->_operator = $comparisonOperator;
        $this->_threshold = $threshold;
    }

    public static function TryCreate($sheet, $area, $formula)
    {
        $expression = trim($formula === null ? '' : $formula);
        if (str_starts_with($expression, '=')) {
            $expression = substr($expression, 1);
        }

        $match = array();
        $threshold = 0.0;
        if (preg_match(self::Comparison, $expression, $match) !== 1
            || !self::TryParseDouble($match[3], $threshold)) {
            return null;
        }

        $token = $match[1];
        $absoluteColumn = $token[0] === '$';
        $rowMarker = strrpos($token, '$');
        $absoluteRow = $rowMarker > 0;
        $reference = null;
        try {
            $reference = CellAddress::Parse(str_replace('$', '', $token));
        } catch (\InvalidArgumentException $ex) {
            return null;
        }

        return new ConditionalFormatExpressionMatcher($sheet, $area, $reference,
            $absoluteRow, $absoluteColumn, $match[2], $threshold);
    }

    private static function TryParseDouble($text, &$value)
    {
        $value = 0.0;
        if (!is_numeric($text)) {
            return false;
        }

        $value = (float)$text;
        return true;
    }

    public function Matches($row, $column)
    {
        $referencedRow = $this->_absoluteRow ? $this->_referenceRow
            : $this->_referenceRow + $row - $this->_anchorRow;
        $referencedColumn = $this->_absoluteColumn ? $this->_referenceColumn
            : $this->_referenceColumn + $column - $this->_anchorColumn;
        if ($referencedRow < 0 || $referencedRow > 2147483647
            || $referencedColumn < 0 || $referencedColumn > 2147483647) {
            return false;
        }

        $key = (new CellAddress((int)$referencedRow, (int)$referencedColumn))->ToString();
        $record = array_key_exists($key, $this->_sheet->Cells) ? $this->_sheet->Cells[$key] : null;
        $value = 0.0;
        if (!ConditionalFormatEvaluator::TryNumber($record, $value)) {
            return false;
        }

        switch ($this->_operator) {
            case '>': return $value > $this->_threshold;
            case '>=': return $value >= $this->_threshold;
            case '<': return $value < $this->_threshold;
            case '<=': return $value <= $this->_threshold;
            case '=': return $value == $this->_threshold;
            case '<>': return $value != $this->_threshold;
            default: return false;
        }
    }
}
