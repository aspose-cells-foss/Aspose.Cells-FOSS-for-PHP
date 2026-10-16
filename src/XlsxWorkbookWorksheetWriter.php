<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellRecord;
use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\ColumnRangeModel;
use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\RowModel;
use Aspose\Cells_FOSS\Core\SharedStringRepository;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

final class XlsxWorkbookWorksheetWriter
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    private function __construct()
    {
    }

    public static function BuildWorksheet(WorksheetModel $worksheet, StyleValue $workbookDefaultStyle, $dateSystem, SharedStringRepository $sharedStrings, SaveOptions $options, StylesheetSaveContext $stylesheet, $externalHyperlinkCount, $hasPictures, $hasComments)
    {
        $persistedCells = self::CollectPersistedCells($worksheet, $workbookDefaultStyle);
        $worksheetElement = new XElement(self::MainNs . 'worksheet',
            new XAttribute(XNamespace::Xmlns . 'r', self::RelationshipNs));

        $sheetProperties = XlsxWorkbookWorksheetViews::BuildWorksheetSheetProperties($worksheet);
        if ($sheetProperties !== null) {
            $worksheetElement->Add($sheetProperties);
        }

        $dimensionReference = self::CalculateDimensionReference($persistedCells, $worksheet->MergeRegions);
        if ($dimensionReference !== '') {
            $worksheetElement->Add(new XElement(self::MainNs . 'dimension', new XAttribute('ref', $dimensionReference)));
        }

        $sheetViews = XlsxWorkbookWorksheetViews::BuildWorksheetViewsElement($worksheet);
        if ($sheetViews !== null) {
            $worksheetElement->Add($sheetViews);
        }

        $normalizedColumns = self::NormalizeColumnRanges($worksheet->Columns);
        if (count($normalizedColumns) > 0) {
            $worksheetElement->Add(new XElement(self::MainNs . 'cols', self::BuildColumnElements($normalizedColumns)));
        }

        $sheetData = new XElement(self::MainNs . 'sheetData');
        $rowIndexes = self::GetWorksheetRowIndexes($persistedCells, $worksheet->Rows);
        for ($i = 0; $i < count($rowIndexes); $i++) {
            $rowIndex = $rowIndexes[$i];
            $rowCells = self::GetRowCells($persistedCells, $rowIndex);
            $rowModel = null;
            if (array_key_exists($rowIndex, $worksheet->Rows)) {
                $rowModel = $worksheet->Rows[$rowIndex];
            }
            if (count($rowCells) === 0 && !self::HasRowMetadata($rowModel)) {
                continue;
            }
            $sheetData->Add(self::BuildRowElement($rowIndex, $rowModel, $rowCells, $dateSystem, $sharedStrings, $options, $stylesheet));
        }

        $worksheetElement->Add($sheetData);

        $sheetProtection = XlsxWorkbookWorksheetProtection::BuildSheetProtectionElement($worksheet);
        if ($sheetProtection !== null) {
            $worksheetElement->Add($sheetProtection);
        }

        $autoFilter = XlsxWorkbookAutoFilter::BuildAutoFilterElement($worksheet, $stylesheet->DifferentialFormatCount);
        if ($autoFilter !== null) {
            $worksheetElement->Add($autoFilter);
        }

        $mergeRegions = self::SortMergeRegions($worksheet->MergeRegions);
        if (count($mergeRegions) > 0) {
            $worksheetElement->Add(new XElement(self::MainNs . 'mergeCells',
                new XAttribute('count', count($mergeRegions)),
                self::BuildMergeCellElements($mergeRegions)));
        }

        $conditionalFormatting = XlsxWorkbookConditionalFormatting::BuildConditionalFormattingElements($worksheet, $stylesheet);
        $worksheetElement->Add($conditionalFormatting);

        $dataValidations = XlsxWorkbookValidations::BuildDataValidationsElement($worksheet);
        if ($dataValidations !== null) {
            $worksheetElement->Add($dataValidations);
        }

        $hyperlinks = XlsxWorkbookHyperlinks::BuildHyperlinksElement($worksheet);
        if ($hyperlinks !== null) {
            $worksheetElement->Add($hyperlinks);
        }

        $printOptions = XlsxWorkbookPageSetup::BuildPrintOptionsElement($worksheet->PageSetup);
        if ($printOptions !== null) {
            $worksheetElement->Add($printOptions);
        }

        $pageMargins = XlsxWorkbookPageSetup::BuildPageMarginsElement($worksheet->PageSetup);
        if ($pageMargins !== null) {
            $worksheetElement->Add($pageMargins);
        }

        $pageSetup = XlsxWorkbookPageSetup::BuildPageSetupElement($worksheet->PageSetup);
        if ($pageSetup !== null) {
            $worksheetElement->Add($pageSetup);
        }

        $headerFooter = XlsxWorkbookPageSetup::BuildHeaderFooterElement($worksheet->PageSetup);
        if ($headerFooter !== null) {
            $worksheetElement->Add($headerFooter);
        }

        $rowBreaks = XlsxWorkbookPageSetup::BuildRowBreaksElement($worksheet->PageSetup);
        if ($rowBreaks !== null) {
            $worksheetElement->Add($rowBreaks);
        }

        $columnBreaks = XlsxWorkbookPageSetup::BuildColumnBreaksElement($worksheet->PageSetup);
        if ($columnBreaks !== null) {
            $worksheetElement->Add($columnBreaks);
        }

        if ($hasPictures) {
            $drawingRId = $externalHyperlinkCount + count($worksheet->ListObjects) + 1;
            $worksheetElement->Add(new XElement(self::MainNs . 'drawing',
                new XAttribute(self::RelationshipNs . 'id', 'rId' . (string)$drawingRId)));
        }

        if ($hasComments) {
            $vmlRId = $externalHyperlinkCount + count($worksheet->ListObjects) + ($hasPictures ? 1 : 0) + 2;
            $worksheetElement->Add(new XElement(self::MainNs . 'legacyDrawing',
                new XAttribute(self::RelationshipNs . 'id', 'rId' . (string)$vmlRId)));
        }

        $tableRIdStart = $externalHyperlinkCount + 1;
        $tableParts = XlsxWorkbookTables::BuildTablePartsElement($worksheet, $tableRIdStart);
        if ($tableParts !== null) {
            $worksheetElement->Add($tableParts);
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $worksheetElement);
    }

    private static function CollectPersistedCells(WorksheetModel $worksheet, StyleValue $workbookDefaultStyle)
    {
        $persistedCells = array();
        foreach ($worksheet->Cells as $addressKey => $record) {
            $addr = CellAddress::Parse($addressKey);
            if (XlsxWorkbookSerializerCommon::ShouldPersistCell($workbookDefaultStyle, $record)) {
                $persistedCells[] = array($addr, $record);
            }
        }
        usort($persistedCells, array(__CLASS__, 'ComparePersistedCells'));
        return $persistedCells;
    }

    private static function ComparePersistedCells($left, $right)
    {
        $rowComparison = $left[0]->RowIndex <=> $right[0]->RowIndex;
        if ($rowComparison !== 0) {
            return $rowComparison;
        }
        return $left[0]->ColumnIndex <=> $right[0]->ColumnIndex;
    }

    private static function BuildColumnElements($columns)
    {
        $columnElements = array();
        for ($index = 0; $index < count($columns); $index++) {
            $columnElements[] = self::BuildColumnElement($columns[$index]);
        }
        return $columnElements;
    }

    private static function GetRowCells($persistedCells, $rowIndex)
    {
        $rowCells = array();
        for ($index = 0; $index < count($persistedCells); $index++) {
            $pair = $persistedCells[$index];
            if ($pair[0]->RowIndex === $rowIndex) {
                $rowCells[] = $pair;
            }
        }
        return $rowCells;
    }

    private static function SortMergeRegions($mergeRegions)
    {
        $sortedRegions = array();
        for ($index = 0; $index < count($mergeRegions); $index++) {
            $sortedRegions[] = $mergeRegions[$index];
        }
        usort($sortedRegions, array(__CLASS__, 'CompareMergeRegions'));
        return $sortedRegions;
    }

    private static function CompareMergeRegions($left, $right)
    {
        $rowComparison = $left->FirstRow <=> $right->FirstRow;
        if ($rowComparison !== 0) {
            return $rowComparison;
        }
        $columnComparison = $left->FirstColumn <=> $right->FirstColumn;
        if ($columnComparison !== 0) {
            return $columnComparison;
        }
        $rowCountComparison = $left->TotalRows <=> $right->TotalRows;
        if ($rowCountComparison !== 0) {
            return $rowCountComparison;
        }
        return $left->TotalColumns <=> $right->TotalColumns;
    }

    private static function BuildMergeCellElements($mergeRegions)
    {
        $mergeCellElements = array();
        for ($index = 0; $index < count($mergeRegions); $index++) {
            $mergeCellElements[] = new XElement(self::MainNs . 'mergeCell', new XAttribute('ref', XlsxWorkbookSerializerCommon::ToRangeReference($mergeRegions[$index])));
        }
        return $mergeCellElements;
    }

    public static function BuildRowElement($rowIndex, $rowModel, $rowCells, $dateSystem, SharedStringRepository $sharedStrings, SaveOptions $options, StylesheetSaveContext $stylesheet)
    {
        $row = new XElement(self::MainNs . 'row', new XAttribute('r', $rowIndex + 1));
        if ($rowModel !== null && $rowModel->Height !== null) {
            $row->SetAttributeValue('ht', (string)$rowModel->Height);
            $row->SetAttributeValue('customHeight', 1);
        }
        if ($rowModel !== null && $rowModel->Hidden) {
            $row->SetAttributeValue('hidden', 1);
        }
        if ($rowModel !== null && $rowModel->StyleIndex !== null && $rowModel->StyleIndex >= 0) {
            $row->SetAttributeValue('s', $rowModel->StyleIndex);
            $row->SetAttributeValue('customFormat', 1);
        }
        foreach ($rowCells as $pair) {
            $row->Add(self::BuildCell($pair[0], $pair[1], $dateSystem, $sharedStrings, $options, $stylesheet));
        }
        return $row;
    }

    public static function BuildColumnElement(ColumnRangeModel $column)
    {
        $element = new XElement(self::MainNs . 'col',
            new XAttribute('min', $column->MinColumnIndex + 1),
            new XAttribute('max', $column->MaxColumnIndex + 1));
        if ($column->Width !== null) {
            $element->SetAttributeValue('width', (string)$column->Width);
            $element->SetAttributeValue('customWidth', 1);
        }
        if ($column->Hidden) {
            $element->SetAttributeValue('hidden', 1);
        }
        if ($column->StyleIndex !== null && $column->StyleIndex >= 0) {
            $element->SetAttributeValue('style', $column->StyleIndex);
        }
        return $element;
    }

    public static function HasRowMetadata($rowModel)
    {
        return $rowModel !== null && ($rowModel->Height !== null || $rowModel->Hidden || $rowModel->StyleIndex !== null);
    }

    public static function HasColumnMetadata(ColumnRangeModel $column)
    {
        return $column->Width !== null || $column->Hidden || $column->StyleIndex !== null;
    }

    public static function GetWorksheetRowIndexes($persistedCells, $rows)
    {
        $indexes = array();
        for ($i = 0; $i < count($persistedCells); $i++) {
            $pair = $persistedCells[$i];
            $idx = $pair[0]->RowIndex;
            $indexes[$idx] = true;
        }
        foreach ($rows as $rowIndex => $rowModel) {
            $indexes[$rowIndex] = true;
        }
        $orderedIndexes = array_keys($indexes);
        sort($orderedIndexes, SORT_NUMERIC);
        return $orderedIndexes;
    }

    public static function NormalizeColumnRanges($columns)
    {
        $ordered = array();
        for ($index = 0; $index < count($columns); $index++) {
            $column = $columns[$index];
            if (!self::HasColumnMetadata($column)) {
                continue;
            }
            $newCol = new ColumnRangeModel();
            $newCol->MinColumnIndex = $column->MinColumnIndex;
            $newCol->MaxColumnIndex = $column->MaxColumnIndex;
            $newCol->Width = $column->Width;
            $newCol->Hidden = $column->Hidden;
            $newCol->StyleIndex = $column->StyleIndex;
            $ordered[] = $newCol;
        }
        usort($ordered, array(__CLASS__, 'CompareColumnRangesByBounds'));
        if (count($ordered) === 0) {
            return $ordered;
        }
        $normalized = array($ordered[0]);
        for ($index = 1; $index < count($ordered); $index++) {
            $current = $ordered[$index];
            $previous = &$normalized[count($normalized) - 1];
            if ($previous->MaxColumnIndex + 1 >= $current->MinColumnIndex && self::ColumnMetadataEqual($previous, $current)) {
                if ($current->MaxColumnIndex > $previous->MaxColumnIndex) {
                    $previous->MaxColumnIndex = $current->MaxColumnIndex;
                }
                unset($previous);
                continue;
            }
            unset($previous);
            $normalized[] = $current;
        }
        return $normalized;
    }

    private static function CompareColumnRangesByBounds($left, $right)
    {
        $minComparison = $left->MinColumnIndex <=> $right->MinColumnIndex;
        if ($minComparison !== 0) {
            return $minComparison;
        }
        return $left->MaxColumnIndex <=> $right->MaxColumnIndex;
    }

    public static function ColumnMetadataEqual(ColumnRangeModel $left, ColumnRangeModel $right)
    {
        return $left->Width === $right->Width
            && $left->Hidden === $right->Hidden
            && $left->StyleIndex === $right->StyleIndex;
    }

    public static function CalculateDimensionReference($persistedCells, $mergeRegions)
    {
        $hasCells = count($persistedCells) > 0;
        $hasMerges = count($mergeRegions) > 0;
        if (!$hasCells && !$hasMerges) {
            return '';
        }
        $minRow = 2147483647;
        $minColumn = 2147483647;
        $maxRow = -2147483648;
        $maxColumn = -2147483648;
        for ($i = 0; $i < count($persistedCells); $i++) {
            $pair = $persistedCells[$i];
            if ($pair[0]->RowIndex < $minRow) {
                $minRow = $pair[0]->RowIndex;
            }
            if ($pair[0]->ColumnIndex < $minColumn) {
                $minColumn = $pair[0]->ColumnIndex;
            }
            if ($pair[0]->RowIndex > $maxRow) {
                $maxRow = $pair[0]->RowIndex;
            }
            if ($pair[0]->ColumnIndex > $maxColumn) {
                $maxColumn = $pair[0]->ColumnIndex;
            }
        }
        for ($i = 0; $i < count($mergeRegions); $i++) {
            $region = $mergeRegions[$i];
            if ($region->FirstRow < $minRow) {
                $minRow = $region->FirstRow;
            }
            if ($region->FirstColumn < $minColumn) {
                $minColumn = $region->FirstColumn;
            }
            $r2 = $region->FirstRow + $region->TotalRows - 1;
            $c2 = $region->FirstColumn + $region->TotalColumns - 1;
            if ($r2 > $maxRow) {
                $maxRow = $r2;
            }
            if ($c2 > $maxColumn) {
                $maxColumn = $c2;
            }
        }
        $firstAddress = (new CellAddress($minRow, $minColumn))->ToString();
        $lastAddress = (new CellAddress($maxRow, $maxColumn))->ToString();
        if ($firstAddress === $lastAddress) {
            return $firstAddress;
        }
        return $firstAddress . ':' . $lastAddress;
    }

    public static function BuildCell(CellAddress $address, CellRecord $record, $dateSystem, SharedStringRepository $sharedStrings, SaveOptions $options, StylesheetSaveContext $stylesheet)
    {
        $cell = new XElement(self::MainNs . 'c', new XAttribute('r', $address->ToString()));
        $styleIndex = $stylesheet->GetStyleIndex($record);
        if ($styleIndex > 0) {
            $cell->SetAttributeValue('s', $styleIndex);
        }
        $hasFormula = !($record->Formula === null || $record->Formula === '');
        if ($hasFormula) {
            $cell->Add(new XElement(self::MainNs . 'f', $record->Formula));
        }
        if ($record->Value === null) {
            return $cell;
        }
        if ($record->Kind === CellValueKind::Error) {
            $cell->SetAttributeValue('t', 'e');
            $cell->Add(new XElement(self::MainNs . 'v', (string)$record->Value));
            return $cell;
        }
        if (is_string($record->Value)) {
            $text = $record->Value;
            if ($hasFormula) {
                $cell->SetAttributeValue('t', 'str');
                $cell->Add(new XElement(self::MainNs . 'v', $text));
            } elseif ($options->UseSharedStrings) {
                $cell->SetAttributeValue('t', 's');
                $cell->Add(new XElement(self::MainNs . 'v', (string)$sharedStrings->Intern($text)));
            } else {
                $cell->SetAttributeValue('t', 'inlineStr');
                $cell->Add(new XElement(self::MainNs . 'is', XlsxWorkbookSerializerCommon::CreateTextElement($text)));
            }
        } elseif (is_bool($record->Value)) {
            $booleanValue = $record->Value;
            $cell->SetAttributeValue('t', 'b');
            $cell->Add(new XElement(self::MainNs . 'v', $booleanValue ? '1' : '0'));
        } elseif ($record->Value instanceof \DateTime) {
            $dateTime = $record->Value;
            $serial = DateSerialConverter::ToSerial($dateTime, $dateSystem);
            $cell->Add(new XElement(self::MainNs . 'v', (string)round($serial, 15)));
        } elseif (is_int($record->Value) || is_long($record->Value) || is_float($record->Value) || is_double($record->Value)) {
            $val = $record->Value;
            if (is_int($val) || is_long($val)) {
                $cell->Add(new XElement(self::MainNs . 'v', (string)$val));
            } else {
                $cell->Add(new XElement(self::MainNs . 'v', (string)$val));
            }
        } else {
            if (is_object($record->Value) && method_exists($record->Value, 'ToString')) {
                $textFallback = $record->Value->ToString();
            } else {
                $textFallback = (string)$record->Value;
            }
            if ($hasFormula) {
                $cell->SetAttributeValue('t', 'str');
                $cell->Add(new XElement(self::MainNs . 'v', $textFallback));
            } elseif ($options->UseSharedStrings) {
                $cell->SetAttributeValue('t', 's');
                $cell->Add(new XElement(self::MainNs . 'v', (string)$sharedStrings->Intern($textFallback)));
            } else {
                $cell->SetAttributeValue('t', 'inlineStr');
                $cell->Add(new XElement(self::MainNs . 'is', XlsxWorkbookSerializerCommon::CreateTextElement($textFallback)));
            }
        }
        return $cell;
    }
}
