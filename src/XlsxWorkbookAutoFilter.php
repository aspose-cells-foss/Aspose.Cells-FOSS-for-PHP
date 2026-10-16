<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterColorFilterModel;
use Aspose\Cells_FOSS\Core\AutoFilterCustomFilterModel;
use Aspose\Cells_FOSS\Core\AutoFilterDynamicFilterModel;
use Aspose\Cells_FOSS\Core\AutoFilterSortConditionModel;
use Aspose\Cells_FOSS\Core\AutoFilterSortStateModel;
use Aspose\Cells_FOSS\Core\AutoFilterTop10Model;
use Aspose\Cells_FOSS\Core\FilterColumnModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookAutoFilter
{
    private function __construct()
    {
    }

    public static function BuildAutoFilterElement(WorksheetModel $worksheet, $differentialFormatCount)
    {
        $autoFilter = $worksheet->AutoFilter;
        if (!$autoFilter->HasStoredState() || $autoFilter->Range === null || $autoFilter->Range === '') {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'autoFilter', new XAttribute('ref', $autoFilter->Range));
        $orderedColumns = self::GetOrderedFilterColumns($autoFilter->FilterColumns);
        for ($index = 0; $index < count($orderedColumns); $index++) {
            $filterColumn = $orderedColumns[$index];
            if (!$filterColumn->HasStoredState()) {
                continue;
            }

            $filterColumnElement = self::BuildFilterColumnElement($filterColumn, $differentialFormatCount);
            if ($filterColumnElement !== null) {
                $element->Add($filterColumnElement);
            }
        }

        $sortState = self::BuildSortStateElement($autoFilter->SortState, $differentialFormatCount);
        if ($sortState !== null) {
            $element->Add($sortState);
        }

        return $element;
    }

    public static function LoadAutoFilter(WorksheetModel $worksheetModel, XElement $worksheetRoot, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->AutoFilter->Clear();
        $autoFilterElement = $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'autoFilter');
        if ($autoFilterElement === null) {
            return;
        }

        $reference = (string)$autoFilterElement->Attribute('ref');
        $normalizedRange = null;
        if (!AutoFilterSupport::TryNormalizeRange($reference, $normalizedRange)) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException("The worksheet autoFilter ref '" . $reference . "' is invalid.");
            }

            $issue = new LoadIssue('WS-L010', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter metadata had an invalid range and was dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $reference;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        $worksheetModel->AutoFilter->Range = $normalizedRange;

        $seenColumnIndexes = array();
        foreach ($autoFilterElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'filterColumn') as $filterColumnElement) {
            $column = self::LoadFilterColumn($filterColumnElement, $stylesheet, $diagnostics, $options, $sheetName, $normalizedRange);
            if ($column === null) {
                continue;
            }

            if (isset($seenColumnIndexes[$column->ColumnIndex])) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The worksheet autoFilter contains duplicate filter columns.');
                }

                $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $normalizedRange;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $seenColumnIndexes[$column->ColumnIndex] = true;
            $worksheetModel->AutoFilter->FilterColumns[] = $column;
        }

        usort($worksheetModel->AutoFilter->FilterColumns, function ($left, $right) {
            return AutoFilterSupport::CompareFilterColumns($left, $right);
        });

        $sortStateElement = $autoFilterElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'sortState');
        if ($sortStateElement !== null) {
            self::LoadSortState($worksheetModel->AutoFilter->SortState, $sortStateElement, $stylesheet, $diagnostics, $options, $sheetName, $normalizedRange);
        }
    }

    private static function HasElements($parent, $name)
    {
        foreach ($parent->Elements($name) as $element) {
            return true;
        }
        return false;
    }

    private static function BuildFilterColumnElement(FilterColumnModel $model, $differentialFormatCount)
    {
        $customFilters = self::BuildCustomFiltersElement($model);
        $filters = self::BuildFiltersElement($model->Filters);
        $colorFilter = self::BuildColorFilterElement($model->ColorFilter, $differentialFormatCount);
        $dynamicFilter = self::BuildDynamicFilterElement($model->DynamicFilter);
        $top10 = self::BuildTop10Element($model->Top10);
        if (!$model->HiddenButton
            && $filters === null
            && $customFilters === null
            && $colorFilter === null
            && $dynamicFilter === null
            && $top10 === null) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'filterColumn', new XAttribute('colId', (string)$model->ColumnIndex));
        if ($model->HiddenButton) {
            $element->SetAttributeValue('hiddenButton', 1);
        }

        if ($colorFilter !== null) {
            $element->Add($colorFilter);
        }

        if ($customFilters !== null) {
            $element->Add($customFilters);
        }

        if ($dynamicFilter !== null) {
            $element->Add($dynamicFilter);
        }

        if ($filters !== null) {
            $element->Add($filters);
        }

        if ($top10 !== null) {
            $element->Add($top10);
        }

        return $element;
    }

    private static function BuildFiltersElement($filters)
    {
        if (count($filters) === 0) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'filters');
        for ($index = 0; $index < count($filters); $index++) {
            $element->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'filter', new XAttribute('val', $filters[$index])));
        }

        return $element;
    }

    private static function BuildCustomFiltersElement(FilterColumnModel $model)
    {
        if (count($model->CustomFilters) === 0) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'customFilters');
        if ($model->CustomFiltersAnd) {
            $element->SetAttributeValue('and', 1);
        }

        for ($index = 0; $index < count($model->CustomFilters); $index++) {
            $filter = $model->CustomFilters[$index];
            $customFilter = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'customFilter', new XAttribute('val', $filter->Value));
            if (!($filter->Operator === null || $filter->Operator === '')) {
                $customFilter->SetAttributeValue('operator', $filter->Operator);
            }

            $element->Add($customFilter);
        }

        return $element;
    }

    private static function BuildColorFilterElement(AutoFilterColorFilterModel $model, $differentialFormatCount)
    {
        if (!$model->Enabled || !self::HasValidDifferentialStyleId($model->DifferentialStyleId, $differentialFormatCount)) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'colorFilter');
        $element->SetAttributeValue('dxfId', (string)$model->DifferentialStyleId);

        if ($model->CellColor) {
            $element->SetAttributeValue('cellColor', 1);
        }

        return $element;
    }

    private static function BuildDynamicFilterElement(AutoFilterDynamicFilterModel $model)
    {
        if (!$model->Enabled) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'dynamicFilter');
        if (!($model->Type === null || $model->Type === '')) {
            $element->SetAttributeValue('type', $model->Type);
        }

        if ($model->Value !== null) {
            $element->SetAttributeValue('val', (string)$model->Value);
        }

        if ($model->MaxValue !== null) {
            $element->SetAttributeValue('maxVal', (string)$model->MaxValue);
        }

        return $element;
    }

    private static function BuildTop10Element(AutoFilterTop10Model $model)
    {
        if (!$model->Enabled) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'top10');
        if (!$model->Top) {
            $element->SetAttributeValue('top', 0);
        }

        if ($model->Percent) {
            $element->SetAttributeValue('percent', 1);
        }

        if ($model->Value !== null) {
            $element->SetAttributeValue('val', (string)$model->Value);
        }

        if ($model->FilterValue !== null) {
            $element->SetAttributeValue('filterVal', (string)$model->FilterValue);
        }

        return $element;
    }

    private static function BuildSortStateElement(AutoFilterSortStateModel $model, $differentialFormatCount)
    {
        if (!$model->HasStoredState() || $model->Ref === null || $model->Ref === '') {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sortState', new XAttribute('ref', $model->Ref));
        if ($model->ColumnSort) {
            $element->SetAttributeValue('columnSort', 1);
        }

        if ($model->CaseSensitive) {
            $element->SetAttributeValue('caseSensitive', 1);
        }

        if (!($model->SortMethod === null || $model->SortMethod === '')) {
            $element->SetAttributeValue('sortMethod', $model->SortMethod);
        }

        for ($index = 0; $index < count($model->Conditions); $index++) {
            $condition = $model->Conditions[$index];
            if ($condition->Ref === null || $condition->Ref === '') {
                continue;
            }

            $conditionElement = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sortCondition', new XAttribute('ref', $condition->Ref));
            $sortBy = self::NormalizeSortBy($condition->SortBy);
            if ($condition->Descending) {
                $conditionElement->SetAttributeValue('descending', 1);
            }

            if (!($sortBy === null || $sortBy === '')) {
                $conditionElement->SetAttributeValue('sortBy', $sortBy);
            }

            if (self::AllowsCustomList($sortBy) && !($condition->CustomList === null || $condition->CustomList === '')) {
                $conditionElement->SetAttributeValue('customList', $condition->CustomList);
            }

            if (self::AllowsDifferentialStyle($sortBy) && self::HasValidDifferentialStyleId($condition->DifferentialStyleId, $differentialFormatCount)) {
                $conditionElement->SetAttributeValue('dxfId', (string)$condition->DifferentialStyleId);
            }

            if (self::AllowsIconSort($condition, $sortBy)) {
                $conditionElement->SetAttributeValue('iconSet', $condition->IconSet);
                $conditionElement->SetAttributeValue('iconId', (string)$condition->IconId);
            }

            $element->Add($conditionElement);
        }

        return $element;
    }

    private static function GetOrderedFilterColumns($columns)
    {
        $ordered = array();
        for ($index = 0; $index < count($columns); $index++) {
            $ordered[] = $columns[$index];
        }

        usort($ordered, function ($left, $right) {
            return AutoFilterSupport::CompareFilterColumns($left, $right);
        });
        return $ordered;
    }

    private static function LoadFilterColumn($filterColumnElement, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $columnIndex = XlsxWorkbookArchiveHelpers::ParseIntAttribute($filterColumnElement->Attribute('colId'));
        if ($columnIndex === null || $columnIndex < 0) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid filter column id.');
            }

            $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        $model = new FilterColumnModel();
        $model->ColumnIndex = $columnIndex;
        $model->HiddenButton = self::ParseFilterColumnHiddenButton($filterColumnElement, $diagnostics, $options, $sheetName, $autoFilterRange);
        self::LoadSimpleFilters($model, $filterColumnElement, $diagnostics, $options, $sheetName, $autoFilterRange);
        self::LoadCustomFilters($model, $filterColumnElement, $diagnostics, $options, $sheetName, $autoFilterRange);
        self::LoadColorFilter($model->ColorFilter, $filterColumnElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'colorFilter'), $stylesheet, $diagnostics, $options, $sheetName, $autoFilterRange);
        self::LoadDynamicFilter($model->DynamicFilter, $filterColumnElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'dynamicFilter'), $diagnostics, $options, $sheetName, $autoFilterRange);
        self::LoadTop10($model->Top10, $filterColumnElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'top10'), $diagnostics, $options, $sheetName, $autoFilterRange);
        return $model;
    }

    private static function ParseFilterColumnHiddenButton($filterColumnElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $hiddenButton = false;
        if (self::TryParseBooleanAttribute($filterColumnElement->Attribute('hiddenButton'), false, $hiddenButton)) {
            return $hiddenButton;
        }

        if ($filterColumnElement->Attribute('hiddenButton') !== null) {
            return self::WarnAndReturnFalse($diagnostics, $options, $sheetName, $autoFilterRange);
        }

        $showButton = false;
        if (self::TryParseBooleanAttribute($filterColumnElement->Attribute('showButton'), true, $showButton)) {
            return !$showButton;
        }

        if ($filterColumnElement->Attribute('showButton') !== null) {
            return self::WarnAndReturnFalse($diagnostics, $options, $sheetName, $autoFilterRange);
        }

        return false;
    }

    private static function WarnAndReturnFalse(LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        if ($options->StrictMode) {
            throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid filter column attribute.');
        }

        $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
        $issue->SheetName = $sheetName;
        $issue->CellRef = $autoFilterRange;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return false;
    }

    private static function LoadSimpleFilters(FilterColumnModel $model, $filterColumnElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $filtersElement = $filterColumnElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'filters');
        if ($filtersElement === null) {
            return;
        }

        if (self::HasElements($filtersElement, XlsxWorkbookSerializerCommon::MainNs . 'dateGroupItem')) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The worksheet autoFilter contains unsupported date group filters.');
            }

            $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        }

        foreach ($filtersElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'filter') as $filterElement) {
            $valueAttribute = $filterElement->Attribute('val');
            $value = (string)$valueAttribute;
            if ($valueAttribute === null) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The worksheet autoFilter contains a filter without a value.');
                }

                $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $autoFilterRange;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $model->Filters[] = $value;
        }
    }

    private static function LoadCustomFilters(FilterColumnModel $model, $filterColumnElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $customFiltersElement = $filterColumnElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'customFilters');
        if ($customFiltersElement === null) {
            return;
        }

        $andValue = false;
        if (self::TryParseBooleanAttribute($customFiltersElement->Attribute('and'), false, $andValue)) {
            $model->CustomFiltersAnd = $andValue;
        } elseif ($customFiltersElement->Attribute('and') !== null) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid custom filter 'and' attribute.");
            }

            $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        }

        foreach ($customFiltersElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'customFilter') as $customFilterElement) {
            $valueAttribute = $customFilterElement->Attribute('val');
            $value = (string)$valueAttribute;
            if ($valueAttribute === null) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The worksheet autoFilter contains a custom filter without a value.');
                }

                $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $autoFilterRange;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $operatorType = FilterOperatorType::Equal;
            if (!AutoFilterSupport::TryParseOperator((string)$customFilterElement->Attribute('operator'), $operatorType)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid custom filter operator.');
                }

                $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $autoFilterRange;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $customFilter = new AutoFilterCustomFilterModel();
            $customFilter->Operator = AutoFilterSupport::ToOperatorName($operatorType) ?? '';
            $customFilter->Value = $value;
            $model->CustomFilters[] = $customFilter;
        }
    }

    private static function LoadColorFilter(AutoFilterColorFilterModel $model, $colorFilterElement, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $model->Clear();
        if ($colorFilterElement === null) {
            return;
        }

        $model->Enabled = true;
        $dxfId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($colorFilterElement->Attribute('dxfId'));
        if ($colorFilterElement->Attribute('dxfId') !== null) {
            if (self::HasValidDifferentialStyleId($dxfId, count($stylesheet->DifferentialFormats))) {
                $model->DifferentialStyleId = $dxfId;
            } else {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid color filter dxfId.');
                }

                $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $autoFilterRange;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                $model->Clear();
                return;
            }
        }

        $cellColor = false;
        if (self::TryParseBooleanAttribute($colorFilterElement->Attribute('cellColor'), false, $cellColor)) {
            $model->CellColor = $cellColor;
        } elseif ($colorFilterElement->Attribute('cellColor') !== null) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid color filter cellColor attribute.');
            }

            $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        }
    }

    private static function LoadDynamicFilter(AutoFilterDynamicFilterModel $model, $dynamicFilterElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $model->Clear();
        if ($dynamicFilterElement === null) {
            return;
        }

        $model->Enabled = true;
        $model->Type = AutoFilterSupport::NormalizeOptionalText((string)$dynamicFilterElement->Attribute('type'));
        $model->Value = self::ParseOptionalDouble($dynamicFilterElement->Attribute('val'), $diagnostics, $options, $sheetName, $autoFilterRange, 'dynamic filter value');
        $model->MaxValue = self::ParseOptionalDouble($dynamicFilterElement->Attribute('maxVal'), $diagnostics, $options, $sheetName, $autoFilterRange, 'dynamic filter max value');
    }

    private static function LoadTop10(AutoFilterTop10Model $model, $top10Element, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $model->Clear();
        if ($top10Element === null) {
            return;
        }

        $model->Enabled = true;
        $topValue = false;
        if (self::TryParseBooleanAttribute($top10Element->Attribute('top'), true, $topValue)) {
            $model->Top = $topValue;
        } elseif ($top10Element->Attribute('top') !== null) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid top10 top attribute.');
            }

            $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        }

        $percentValue = false;
        if (self::TryParseBooleanAttribute($top10Element->Attribute('percent'), false, $percentValue)) {
            $model->Percent = $percentValue;
        } elseif ($top10Element->Attribute('percent') !== null) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid top10 percent attribute.');
            }

            $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        }

        $model->Value = self::ParseOptionalDouble($top10Element->Attribute('val'), $diagnostics, $options, $sheetName, $autoFilterRange, 'top10 value');
        $model->FilterValue = self::ParseOptionalDouble($top10Element->Attribute('filterVal'), $diagnostics, $options, $sheetName, $autoFilterRange, 'top10 filter value');
    }

    private static function LoadSortState(AutoFilterSortStateModel $model, $sortStateElement, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $model->Clear();
        $reference = (string)$sortStateElement->Attribute('ref');
        $normalizedRange = null;
        if (!AutoFilterSupport::TryNormalizeRange($reference, $normalizedRange)) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The worksheet autoFilter sortState ref is invalid.');
            }

            $issue = new LoadIssue('WS-L012', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter sort state was invalid and was dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $autoFilterRange;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        $model->Ref = $normalizedRange;
        $columnSort = false;
        if (self::TryParseBooleanAttribute($sortStateElement->Attribute('columnSort'), false, $columnSort)) {
            $model->ColumnSort = $columnSort;
        } elseif ($sortStateElement->Attribute('columnSort') !== null) {
            self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
        }

        $caseSensitive = false;
        if (self::TryParseBooleanAttribute($sortStateElement->Attribute('caseSensitive'), false, $caseSensitive)) {
            $model->CaseSensitive = $caseSensitive;
        } elseif ($sortStateElement->Attribute('caseSensitive') !== null) {
            self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
        }

        $model->SortMethod = AutoFilterSupport::NormalizeOptionalText((string)$sortStateElement->Attribute('sortMethod'));
        foreach ($sortStateElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'sortCondition') as $sortConditionElement) {
            $conditionReference = (string)$sortConditionElement->Attribute('ref');
            $normalizedConditionRef = null;
            if (!AutoFilterSupport::TryNormalizeRange($conditionReference, $normalizedConditionRef)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The worksheet autoFilter sortCondition ref is invalid.');
                }

                self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
                continue;
            }

            $condition = new AutoFilterSortConditionModel();
            $condition->Ref = $normalizedConditionRef;
            $descending = false;
            if (self::TryParseBooleanAttribute($sortConditionElement->Attribute('descending'), false, $descending)) {
                $condition->Descending = $descending;
            } elseif ($sortConditionElement->Attribute('descending') !== null) {
                self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
            }

            $sortBy = self::NormalizeSortBy((string)$sortConditionElement->Attribute('sortBy'));
            if ($sortConditionElement->Attribute('sortBy') !== null && ($sortBy === null || $sortBy === '')) {
                self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
            }

            $condition->SortBy = $sortBy;

            $customList = AutoFilterSupport::NormalizeOptionalText((string)$sortConditionElement->Attribute('customList'));
            if (self::AllowsCustomList($sortBy)) {
                $condition->CustomList = $customList ?? '';
            } elseif (!($customList === null || $customList === '')) {
                self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
            }

            $iconSet = AutoFilterSupport::NormalizeOptionalText((string)$sortConditionElement->Attribute('iconSet'));
            if (self::AllowsIconSort($sortBy, $iconSet)) {
                $condition->IconSet = $iconSet ?? '';
            } elseif (!($iconSet === null || $iconSet === '')) {
                self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
            }

            $dxfId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($sortConditionElement->Attribute('dxfId'));
            if ($sortConditionElement->Attribute('dxfId') !== null) {
                if (self::AllowsDifferentialStyle($sortBy) && self::HasValidDifferentialStyleId($dxfId, count($stylesheet->DifferentialFormats))) {
                    $condition->DifferentialStyleId = $dxfId;
                } else {
                    if ($options->StrictMode) {
                        throw new InvalidFileFormatException('The worksheet autoFilter sortCondition dxfId is invalid.');
                    }

                    self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
                }
            }

            $iconId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($sortConditionElement->Attribute('iconId'));
            if ($sortConditionElement->Attribute('iconId') !== null) {
                if (self::AllowsIconSort($sortBy, $iconSet) && $iconId !== null && $iconId >= 0) {
                    $condition->IconId = $iconId;
                } else {
                    if ($options->StrictMode) {
                        throw new InvalidFileFormatException('The worksheet autoFilter sortCondition iconId is invalid.');
                    }

                    self::WarnSortState($diagnostics, $options, $sheetName, $autoFilterRange);
                }
            }

            $model->Conditions[] = $condition;
        }
    }

    private static function WarnSortState(LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange)
    {
        $issue = new LoadIssue('WS-L012', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter sort state was invalid and was dropped during load.', true, true);
        $issue->SheetName = $sheetName;
        $issue->CellRef = $autoFilterRange;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
    }

    private static function TryParseBooleanAttribute($attribute, $defaultValue, &$parsedValue)
    {
        $parsedValue = $defaultValue;
        if ($attribute === null) {
            return true;
        }

        $rawValue = trim((string)$attribute ?? '');
        if ($rawValue === '1' || strcasecmp($rawValue, 'true') === 0) {
            $parsedValue = true;
            return true;
        }

        if ($rawValue === '0' || strcasecmp($rawValue, 'false') === 0) {
            $parsedValue = false;
            return true;
        }

        return false;
    }

    private static function ParseOptionalDouble($attribute, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $autoFilterRange, $attributeLabel)
    {
        if ($attribute === null) {
            return null;
        }

        if (is_numeric($attribute->Value)) {
            return (float)$attribute->Value;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException('The worksheet autoFilter contains an invalid ' . $attributeLabel . '.');
        }

        $issue = new LoadIssue('WS-L011', DiagnosticSeverity::LossyRecoverable, 'Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.', true, true);
        $issue->SheetName = $sheetName;
        $issue->CellRef = $autoFilterRange;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return null;
    }

    private static function HasValidDifferentialStyleId($differentialStyleId, $differentialFormatCount)
    {
        return $differentialStyleId !== null
            && $differentialStyleId >= 0
            && $differentialStyleId < $differentialFormatCount;
    }

    private static function NormalizeSortBy($sortBy)
    {
        if ($sortBy === null || trim($sortBy) === '') {
            return '';
        }

        $normalized = trim($sortBy);
        if (strcasecmp($normalized, 'value') === 0) {
            return 'value';
        }

        if (strcasecmp($normalized, 'cellColor') === 0) {
            return 'cellColor';
        }

        if (strcasecmp($normalized, 'fontColor') === 0) {
            return 'fontColor';
        }

        if (strcasecmp($normalized, 'icon') === 0) {
            return 'icon';
        }

        return '';
    }

    private static function AllowsCustomList($sortBy)
    {
        return $sortBy === null || $sortBy === '' || $sortBy === 'value';
    }

    private static function AllowsDifferentialStyle($sortBy)
    {
        return $sortBy === 'cellColor'
            || $sortBy === 'fontColor';
    }

    private static function AllowsIconSort($conditionOrSortBy, $sortByOrIconSet)
    {
        if (is_object($conditionOrSortBy)) {
            $condition = $conditionOrSortBy;
            $sortBy = $sortByOrIconSet;
            return $sortBy === 'icon'
                && !($condition->IconSet === null || $condition->IconSet === '')
                && $condition->IconId !== null
                && $condition->IconId >= 0;
        }

        $sortBy = $conditionOrSortBy;
        $iconSet = $sortByOrIconSet;
        return $sortBy === 'icon'
            && !($iconSet === null || $iconSet === '');
    }
}
