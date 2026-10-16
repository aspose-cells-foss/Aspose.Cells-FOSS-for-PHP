<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\FilterColumnModel;

final class AutoFilterSupport
{
    private function __construct()
    {
    }

    public static function NormalizeOptionalRange($value, $parameterName)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        $normalized = '';
        if (!self::TryNormalizeRange($value, $normalized)) {
            throw new CellsException($parameterName . ' must be a valid cell or range reference.');
        }

        return $normalized;
    }

    public static function NormalizeRequiredRange($value, $parameterName)
    {
        if ($value === null || trim($value) === '') {
            throw new CellsException($parameterName . ' must be a valid cell or range reference.');
        }

        $normalized = '';
        if (!self::TryNormalizeRange($value, $normalized)) {
            throw new CellsException($parameterName . ' must be a valid cell or range reference.');
        }

        return $normalized;
    }

    public static function TryNormalizeRange($value, &$normalized)
    {
        $normalized = '';
        if ($value === null || trim($value) === '') {
            return false;
        }

        $region = null;
        if (!XlsxWorkbookSerializerCommon::TryParseMergeReference(trim($value), $region)) {
            return false;
        }

        $normalized = XlsxWorkbookSerializerCommon::ToRangeReference($region);
        return true;
    }

    public static function NormalizeText($value, $parameterName)
    {
        if ($value === null) {
            throw new \InvalidArgumentException($parameterName);
        }

        return $value;
    }

    public static function NormalizeOptionalText($value)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        return trim($value);
    }

    public static function ParseOperatorOrDefault($value)
    {
        if ($value === null || trim($value) === '') {
            return FilterOperatorType::Equal;
        }

        switch (trim($value)) {
            case 'lessThan':
                return FilterOperatorType::LessThan;
            case 'lessThanOrEqual':
                return FilterOperatorType::LessOrEqual;
            case 'notEqual':
                return FilterOperatorType::NotEqual;
            case 'greaterThanOrEqual':
                return FilterOperatorType::GreaterOrEqual;
            case 'greaterThan':
                return FilterOperatorType::GreaterThan;
            default:
                return FilterOperatorType::Equal;
        }
    }

    public static function TryParseOperator($value, &$operatorType)
    {
        $operatorType = FilterOperatorType::Equal;
        if ($value === null || trim($value) === '') {
            return true;
        }

        switch (trim($value)) {
            case 'lessThan':
                $operatorType = FilterOperatorType::LessThan;
                return true;
            case 'lessThanOrEqual':
                $operatorType = FilterOperatorType::LessOrEqual;
                return true;
            case 'notEqual':
                $operatorType = FilterOperatorType::NotEqual;
                return true;
            case 'greaterThanOrEqual':
                $operatorType = FilterOperatorType::GreaterOrEqual;
                return true;
            case 'greaterThan':
                $operatorType = FilterOperatorType::GreaterThan;
                return true;
            case 'equal':
                $operatorType = FilterOperatorType::Equal;
                return true;
            default:
                return false;
        }
    }

    public static function ToOperatorName($operatorType)
    {
        switch ($operatorType) {
            case FilterOperatorType::LessThan:
                return 'lessThan';
            case FilterOperatorType::LessOrEqual:
                return 'lessThanOrEqual';
            case FilterOperatorType::NotEqual:
                return 'notEqual';
            case FilterOperatorType::GreaterOrEqual:
                return 'greaterThanOrEqual';
            case FilterOperatorType::GreaterThan:
                return 'greaterThan';
            default:
                return null;
        }
    }

    public static function CompareFilterColumns(FilterColumnModel $left, FilterColumnModel $right)
    {
        if ($left->ColumnIndex < $right->ColumnIndex) {
            return -1;
        }
        if ($left->ColumnIndex > $right->ColumnIndex) {
            return 1;
        }
        return 0;
    }
}
