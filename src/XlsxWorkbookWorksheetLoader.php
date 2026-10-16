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
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

final class XlsxWorkbookWorksheetLoader
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';

    private function __construct()
    {
    }

    public static function LoadWorksheet(
        $sheetName,
        $worksheetDocument,
        $worksheetHyperlinkTargets,
        $dateSystem,
        $sharedStrings,
        StylesheetLoadContext $stylesheet,
        LoadDiagnostics $diagnostics,
        LoadOptions $options,
        $definedNamesState,
        ZipArchive $archive,
        $worksheetUri)
    {
        $worksheetModel = new WorksheetModel($sheetName);
        $worksheetRoot = null;
        if ($worksheetDocument instanceof \Aspose\Cells_FOSS\Internal\Xml\XDocument) {
            $worksheetRoot = $worksheetDocument->Root;
        }
        if ($worksheetRoot === null) {
            return $worksheetModel;
        }

        XlsxWorkbookWorksheetViews::LoadWorksheetViewSettings($worksheetModel, $worksheetRoot, $diagnostics, $options, $sheetName);
        XlsxWorkbookWorksheetProtection::LoadWorksheetProtection($worksheetModel, $worksheetRoot, $diagnostics, $options, $sheetName);
        XlsxWorkbookAutoFilter::LoadAutoFilter($worksheetModel, $worksheetRoot, $stylesheet, $diagnostics, $options, $sheetName);
        self::LoadColumns($worksheetModel, $worksheetRoot, $stylesheet, $diagnostics, $options, $sheetName);
        XlsxWorkbookPageSetup::LoadWorksheetPageSetup($worksheetModel, $worksheetRoot, $diagnostics, $options, $sheetName);
        XlsxWorkbookHyperlinks::LoadHyperlinks($worksheetModel, $worksheetRoot, $worksheetHyperlinkTargets, $diagnostics, $options, $sheetName);
        XlsxWorkbookConditionalFormatting::LoadConditionalFormattings($worksheetModel, $worksheetRoot, $stylesheet, $diagnostics, $options, $sheetName);
        XlsxWorkbookValidations::LoadValidations($worksheetModel, $worksheetRoot, $diagnostics, $options, $sheetName);
        XlsxWorkbookTables::LoadTables($worksheetModel, $worksheetRoot, $archive, $worksheetUri, $diagnostics, $options, $sheetName);
        XlsxWorkbookPicturesLoader::LoadPictures($worksheetModel, $worksheetRoot, $archive, $worksheetUri, $diagnostics, $options, $sheetName);
        XlsxWorkbookCommentsLoader::LoadComments($worksheetModel, $archive, $worksheetUri, $diagnostics, $options, $sheetName);
            XlsxWorkbookPageSetup::ApplyWorksheetDefinedNames($worksheetModel, $definedNamesState);

        $sheetData = $worksheetRoot->Element(self::MainNs . 'sheetData');
        if ($sheetData === null) {
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('ACF-WS-001', DiagnosticSeverity::Recoverable, 'Worksheet sheetData is missing; an empty sheet was synthesized.', true));
            return $worksheetModel;
        }

        $rowElements = $sheetData->Elements(self::MainNs . 'row');
        $seenRows = array();
        $seenCells = array();
        $previousRowIndex = -1;
        $rowsOutOfOrderReported = false;

        foreach ($rowElements as $rowElement) {
            $rowIndex = 0;
            if (!self::TryResolveRowIndex($rowElement, $diagnostics, $options, $sheetName, $rowIndex)) {
                continue;
            }
            if (array_key_exists($rowIndex, $seenRows)) {
                throw new InvalidFileFormatException('Duplicate row index was found in worksheet.');
            }
            $seenRows[$rowIndex] = true;
            if ($previousRowIndex > $rowIndex && !$rowsOutOfOrderReported) {
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('WS-R002', DiagnosticSeverity::Recoverable, 'Worksheet rows were out of order and were normalized during load.', true));
                $rowsOutOfOrderReported = true;
            }
            $previousRowIndex = $rowIndex;
            self::ApplyRowMetadata($worksheetModel, $rowElement, $rowIndex, $stylesheet, $diagnostics, $options, $sheetName);

            $previousColumnIndex = -1;
            $cellsOutOfOrderReported = false;
            foreach ($rowElement->Elements(self::MainNs . 'c') as $cellElement) {
                $address = null;
                $record = null;
                if (!self::TryReadCellRecord($cellElement, $rowIndex, $dateSystem, $sharedStrings, $stylesheet, $diagnostics, $options, $sheetName, $address, $record)) {
                    continue;
                }
                if ($previousColumnIndex > $address->ColumnIndex && !$cellsOutOfOrderReported) {
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('WS-R003', DiagnosticSeverity::Recoverable, 'Worksheet cells were out of order within a row and were normalized during load.', true));
                    $cellsOutOfOrderReported = true;
                }
                $previousColumnIndex = $address->ColumnIndex;
                $addrKey = $address->ToString();
                if (array_key_exists($addrKey, $seenCells)) {
                    throw new InvalidFileFormatException('Duplicate cell reference was found in worksheet.');
                }
                $seenCells[$addrKey] = true;
                if (XlsxWorkbookSerializerCommon::ShouldPersistCell($stylesheet->DefaultCellStyle, $record)) {
                    $worksheetModel->Cells[$addrKey] = $record;
                }
            }
        }

        self::LoadMergeRegions($worksheetModel, $worksheetRoot, $diagnostics, $options, $sheetName);

        if ($worksheetRoot->Element(self::MainNs . 'dimension') === null && (count($worksheetModel->Cells) > 0 || count($worksheetModel->MergeRegions) > 0)) {
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('WS-R001', DiagnosticSeverity::Recoverable, 'Worksheet dimension was missing and was recalculated during load.', true));
        }

        return $worksheetModel;
    }

    private static function LoadColumns(WorksheetModel $worksheetModel, $worksheetRoot, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $columns = array();
        $colsElement = $worksheetRoot->Element(self::MainNs . 'cols');
        if ($colsElement !== null) {
            foreach ($colsElement->Elements(self::MainNs . 'col') as $columnElement) {
                $minAttr = $columnElement->Attribute('min');
                $maxAttr = $columnElement->Attribute('max');
                $min = $minAttr !== null ? intval($minAttr->Value) : null;
                $max = $maxAttr !== null ? intval($maxAttr->Value) : null;
                if ($min === null || $max === null || $min <= 0 || $max <= 0 || $min > $max) {
                    throw new InvalidFileFormatException('Worksheet column metadata contains an invalid min/max span.');
                }
                $styleIndex = null;
                $styleAttr = $columnElement->Attribute('style');
                if ($styleAttr !== null) {
                    $styleIndex = intval($styleAttr->Value);
                    if ($styleIndex < 0 || $styleIndex >= count($stylesheet->CellFormats)) {
                        if ($options->StrictMode) {
                            throw new InvalidFileFormatException('The column style index is invalid.');
                        }
                        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('COL-L001', DiagnosticSeverity::Warning, 'Column style index is invalid and was dropped.', false, true));
                        $styleIndex = null;
                    }
                }
                $widthAttr = $columnElement->Attribute('width');
                $width = null;
                if ($widthAttr !== null && $widthAttr->Value !== '') {
                    $width = floatval($widthAttr->Value);
                }
                $hidden = false;
                $hiddenAttr = $columnElement->Attribute('hidden');
                if ($hiddenAttr !== null && $hiddenAttr->Value !== '' && intval($hiddenAttr->Value) !== 0) {
                    $hidden = true;
                }
                $colModel = new ColumnRangeModel();
                $colModel->MinColumnIndex = $min - 1;
                $colModel->MaxColumnIndex = $max - 1;
                $colModel->Width = $width;
                $colModel->Hidden = $hidden;
                $colModel->StyleIndex = $styleIndex;
                $columns[] = $colModel;
            }
        }
        $normalized = self::NormalizeLoadedColumns($columns, $diagnostics, $options, $sheetName);
        foreach ($normalized as $n) {
            $worksheetModel->Columns[] = $n;
        }
    }

    private static function NormalizeLoadedColumns($columns, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $ordered = array();
        for ($i = 0; $i < count($columns); $i++) {
            $ordered[] = $columns[$i];
        }
        usort($ordered, array(__CLASS__, 'CompareColumnRangesByBounds'));
        if (count($ordered) === 0) {
            return $ordered;
        }
        $normalized = array($ordered[0]);
        $mergeReported = false;
        for ($i = 1; $i < count($ordered); $i++) {
            $current = $ordered[$i];
            $previous = &$normalized[count($normalized) - 1];
            if ($current->MinColumnIndex <= $previous->MaxColumnIndex + 1 && self::ColumnRangesCompatible($previous, $current)) {
                if (($current->MinColumnIndex <= $previous->MaxColumnIndex || $current->MaxColumnIndex <= $previous->MaxColumnIndex) && !$mergeReported) {
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('COL-R001', DiagnosticSeverity::Recoverable, 'Overlapping compatible column metadata was normalized during load.', true));
                    $mergeReported = true;
                }
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

    private static function ColumnRangesCompatible($left, $right)
    {
        return $left->Width === $right->Width && $left->Hidden === $right->Hidden && $left->StyleIndex === $right->StyleIndex;
    }

    private static function TryResolveRowIndex($rowElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, &$rowIndex)
    {
        $rowIndexAttr = $rowElement->Attribute('r');
        if ($rowIndexAttr !== null && $rowIndexAttr->Value !== '') {
            $val = intval($rowIndexAttr->Value);
            if ($val <= 0) {
                throw new InvalidFileFormatException('Worksheet row index is invalid.');
            }
            $rowIndex = $val - 1;
            return true;
        }
        foreach ($rowElement->Elements(self::MainNs . 'c') as $cellElement) {
            $cellRef = $cellElement->Attribute('r');
            if ($cellRef !== null && $cellRef->Value !== '') {
                $addr = null;
                if (XlsxWorkbookSerializerCommon::TryParseCellReference($cellRef->Value, $addr)) {
                    $rowIndex = $addr->RowIndex;
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('WS-R004', DiagnosticSeverity::Recoverable, 'A worksheet row index was missing and was inferred from contained cells.', true));
                    return true;
                }
            }
        }
        $rowIndex = -1;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('ROW-F001', DiagnosticSeverity::Warning, 'A worksheet row without an index and without parseable cells was skipped.'));
        return false;
    }

    private static function ApplyRowMetadata(WorksheetModel $worksheetModel, $rowElement, $rowIndex, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $styleIndex = null;
        $styleAttr = $rowElement->Attribute('s');
        if ($styleAttr !== null && $styleAttr->Value !== '') {
            $styleIndex = intval($styleAttr->Value);
            if ($styleIndex < 0 || $styleIndex >= count($stylesheet->CellFormats)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException('The row style index is invalid.');
                }
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('ROW-L001', DiagnosticSeverity::Warning, 'Row style index is invalid and was dropped.', false, true));
                $styleIndex = null;
            }
        }
        $height = null;
        $htAttr = $rowElement->Attribute('ht');
        if ($htAttr !== null && $htAttr->Value !== '') {
            $height = floatval($htAttr->Value);
        }
        $hidden = false;
        $hiddenAttr = $rowElement->Attribute('hidden');
        if ($hiddenAttr !== null && $hiddenAttr->Value !== '' && intval($hiddenAttr->Value) !== 0) {
            $hidden = true;
        }
        $rowModel = new RowModel();
        $rowModel->Height = $height;
        $rowModel->Hidden = $hidden;
        $rowModel->StyleIndex = $styleIndex;
        if ($rowModel->Height !== null) {
            $customHtAttr = $rowElement->Attribute('customHeight');
            if ($customHtAttr === null || $customHtAttr->Value === '' || intval($customHtAttr->Value) === 0) {
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('ROW-R002', DiagnosticSeverity::Recoverable, 'Row height metadata was missing customHeight and was normalized during load.', true));
            }
        }
        if ($rowModel->Height !== null || $rowModel->Hidden || $rowModel->StyleIndex !== null) {
            $worksheetModel->Rows[$rowIndex] = $rowModel;
        }
    }

    private static function TryReadCellRecord($cellElement, $rowIndex, $dateSystem, $sharedStrings, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, &$address, &$record)
    {
        $record = new CellRecord();
        $record->Style = $stylesheet->DefaultCellStyle->Clone();
        $record->IsExplicitlyStored = true;
        $cellRefAttr = $cellElement->Attribute('r');
        if ($cellRefAttr === null || trim($cellRefAttr->Value) === '') {
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('CELL-F001', DiagnosticSeverity::Warning, 'A cell without a reference was skipped.'));
            $address = null;
            return false;
        }
        $resolvedCellReference = $cellRefAttr->Value;
        try {
            $address = CellAddress::Parse($resolvedCellReference);
        } catch (\Exception $e) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('The cell reference is invalid.');
            }
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('CELL-F001', DiagnosticSeverity::Warning, 'Cell reference is invalid and was skipped.'));
            $address = null;
            return false;
        }
        $styleAttr = $cellElement->Attribute('s');
        $styleIndex = null;
        if ($styleAttr !== null && $styleAttr->Value !== '') {
            $styleIndex = intval($styleAttr->Value);
        }
        $isDateStyle = false;
        if ($styleIndex !== null && isset($stylesheet->DateStyleIndexes[$styleIndex])) {
            $isDateStyle = true;
        }
        if ($styleIndex !== null) {
            if ($styleIndex >= 0 && $styleIndex < count($stylesheet->CellFormats)) {
                $record->Style = $stylesheet->CellFormats[$styleIndex]->Clone();
            } elseif ($options->StrictMode) {
                throw new InvalidFileFormatException('The style index is invalid.');
            } else {
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('STYLE-F001', DiagnosticSeverity::Warning, 'Cell style index is invalid and style 0 was used instead.'));
            }
        }
        $formulaElement = $cellElement->Element(self::MainNs . 'f');
        $formulaText = '';
        if ($formulaElement !== null) {
            $formulaText = XlsxWorkbookArchiveHelpers::NormalizeFormula($formulaElement->Value);
        }
        if ($formulaText !== '') {
            $record->Formula = $formulaText;
            $record->Kind = CellValueKind::Formula;
        }
        $cellType = '';
        $cellTypeAttr = $cellElement->Attribute('t');
        if ($cellTypeAttr !== null) {
            $cellType = $cellTypeAttr->Value;
        }
        $valueElement = $cellElement->Element(self::MainNs . 'v');
        $rawValue = null;
        if ($valueElement !== null) {
            $rawValue = $valueElement->Value;
        }
        $value = null;
        $kind = CellValueKind::Blank;
        if (self::TryReadCellValue($cellElement, $cellType, $rawValue, $isDateStyle, $dateSystem, $sharedStrings, $diagnostics, $options, $sheetName, $resolvedCellReference, $value, $kind)) {
            $record->Value = $value;
            if ($record->Formula === null || $record->Formula === '' || $kind === CellValueKind::Error) {
                $record->Kind = $kind;
            }
        }
        return true;
    }

    private static function LoadMergeRegions(WorksheetModel $worksheetModel, $worksheetRoot, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $mergeCellsElement = $worksheetRoot->Element(self::MainNs . 'mergeCells');
        if ($mergeCellsElement !== null) {
            foreach ($mergeCellsElement->Elements(self::MainNs . 'mergeCell') as $mergeElement) {
                $mergeRefAttr = $mergeElement->Attribute('ref');
                $mergeReference = $mergeRefAttr !== null ? $mergeRefAttr->Value : '';
                $region = null;
                if (!XlsxWorkbookSerializerCommon::TryParseMergeReference($mergeReference, $region)) {
                    throw new InvalidFileFormatException('The merge reference is invalid.');
                }
                if (self::ContainsOverlappingMergeRegion($worksheetModel->MergeRegions, $region)) {
                    if ($options->StrictMode) {
                        throw new InvalidFileFormatException('The merge reference overlaps an existing merged range.');
                    }
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('MRG-L001', DiagnosticSeverity::LossyRecoverable, 'Overlapping merge range was dropped during load.', true, true));
                    continue;
                }
                $worksheetModel->MergeRegions[] = $region;
            }
        }
        usort($worksheetModel->MergeRegions, function ($left, $right) {
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
        });
    }

    private static function ContainsOverlappingMergeRegion($mergeRegions, MergeRegion $candidate)
    {
        for ($i = 0; $i < count($mergeRegions); $i++) {
            if (self::MergeRegionsOverlap($mergeRegions[$i], $candidate)) {
                return true;
            }
        }
        return false;
    }

    private static function MergeRegionsOverlap(MergeRegion $left, MergeRegion $right)
    {
        $leftLastRow = $left->FirstRow + $left->TotalRows - 1;
        $leftLastColumn = $left->FirstColumn + $left->TotalColumns - 1;
        $rightLastRow = $right->FirstRow + $right->TotalRows - 1;
        $rightLastColumn = $right->FirstColumn + $right->TotalColumns - 1;
        return $left->FirstRow <= $rightLastRow && $right->FirstRow <= $leftLastRow && $left->FirstColumn <= $rightLastColumn && $right->FirstColumn <= $leftLastColumn;
    }

    private static function TryReadCellValue($cellElement, $cellType, $rawValue, $isDateStyle, $dateSystem, $sharedStrings, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellReference, &$value, &$kind)
    {
        $value = null;
        $kind = CellValueKind::Blank;
        if ($cellType === 'inlineStr') {
            $isElement = $cellElement->Element(self::MainNs . 'is');
            $value = XlsxWorkbookArchiveHelpers::ReadInlineString($isElement);
            $kind = CellValueKind::String;
            return true;
        }
        if ($cellType === 's') {
            $sharedStringIndex = intval($rawValue);
            if (is_array($sharedStrings) || $sharedStrings instanceof \ArrayAccess) {
                if ($sharedStringIndex >= 0 && $sharedStringIndex < count($sharedStrings)) {
                    $value = $sharedStrings[$sharedStringIndex];
                    $kind = CellValueKind::String;
                    return true;
                }
            }
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('SST-L001', DiagnosticSeverity::LossyRecoverable, 'The cell points to an invalid shared string index.', false, true));
            $value = '';
            $kind = CellValueKind::String;
            return true;
        }
        if ($cellType === 'b') {
            $value = ($rawValue === '1' || strtolower($rawValue) === 'true');
            $kind = CellValueKind::Boolean;
            return true;
        }
        if ($cellType === 'd') {
            try {
                $dt = new \DateTime($rawValue);
                $value = $dt;
                $kind = CellValueKind::DateTime;
                return true;
            } catch (\Exception $e) {
                return false;
            }
        }
        if ($cellType === 'str') {
            $value = $rawValue === null ? '' : $rawValue;
            $kind = CellValueKind::String;
            return true;
        }
        if ($cellType === 'e') {
            $value = $rawValue === null ? '' : $rawValue;
            $kind = CellValueKind::Error;
            return true;
        }
        if ($rawValue === null || $rawValue === '') {
            return false;
        }
        if ($isDateStyle) {
            if (is_numeric($rawValue)) {
                $serial = floatval($rawValue);
                $value = DateSerialConverter::FromSerial($serial, $dateSystem);
                $kind = CellValueKind::DateTime;
                return true;
            }
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('CELL-R002', DiagnosticSeverity::Recoverable, 'A formula or numeric cell contained an invalid cached date serial and the cached value was cleared.', true));
            return false;
        }
        if (is_numeric($rawValue)) {
            if (strpos($rawValue, '.') !== false || strpos($rawValue, 'e') !== false || strpos($rawValue, 'E') !== false) {
                $value = floatval($rawValue);
            } else {
                $intval = intval($rawValue);
                if ((string)$intval === $rawValue) {
                    $value = $intval;
                } else {
                    $value = floatval($rawValue);
                }
            }
            $kind = CellValueKind::Number;
            return true;
        }
        return false;
    }
}
