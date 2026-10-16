<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\ListColumnModel;
use Aspose\Cells_FOSS\Core\ListObjectModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

final class XlsxWorkbookTables
{
    const TableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
    const TableContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

    private function __construct()
    {
    }

    public static function BuildTableDocument($model, $tableId)
    {
        $startAddress = (new CellAddress($model->StartRow, $model->StartColumn))->ToString();
        $endAddress = (new CellAddress($model->EndRow, $model->EndColumn))->ToString();
        $tableRef = $startAddress . ':' . $endAddress;

        $tableElement = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'table',
            new XAttribute(XNamespace::Xmlns . 'r', XlsxWorkbookSerializerCommon::RelationshipNs),
            new XAttribute('id', (string)$tableId),
            new XAttribute('name', $model->Name),
            new XAttribute('displayName', $model->DisplayName),
            new XAttribute('ref', $tableRef)
        );

        if (!$model->ShowHeaderRow) {
            $tableElement->SetAttributeValue('headerRowCount', '0');
        }

        if ($model->ShowTotals) {
            $tableElement->SetAttributeValue('totalsRowCount', '1');
        } else {
            $tableElement->SetAttributeValue('totalsRowShown', '0');
        }

        if (!($model->Comment === null || $model->Comment === '')) {
            $tableElement->SetAttributeValue('comment', $model->Comment);
        }

        if ($model->HasAutoFilter && $model->ShowHeaderRow) {
            $autoFilterRef = self::BuildAutoFilterRef($model);
            $tableElement->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'autoFilter', new XAttribute('ref', $autoFilterRef)));
        }

        $tableColumns = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'tableColumns',
            new XAttribute('count', count($model->Columns))
        );

        for ($i = 0; $i < count($model->Columns); $i++) {
            $tableColumns->Add(self::BuildTableColumnElement($model->Columns[$i]));
        }

        $tableElement->Add($tableColumns);

        $tableStyleInfo = self::BuildTableStyleInfoElement($model);
        if ($tableStyleInfo !== null) {
            $tableElement->Add($tableStyleInfo);
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $tableElement);
    }

    private static function BuildAutoFilterRef($model)
    {
        $dataEndRow = $model->ShowTotals ? $model->EndRow - 1 : $model->EndRow;
        $startAddress = (new CellAddress($model->StartRow, $model->StartColumn))->ToString();
        $endAddress = (new CellAddress($dataEndRow, $model->EndColumn))->ToString();
        return $startAddress . ':' . $endAddress;
    }

    private static function BuildTableColumnElement($column)
    {
        $element = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'tableColumn',
            new XAttribute('id', (string)$column->Id),
            new XAttribute('name', $column->Name)
        );

        $function = $column->TotalsRowFunction;
        if (!($function === null || $function === '') && !($function === 'none')) {
            $element->SetAttributeValue('totalsRowFunction', $function);
        }

        if (($function === 'none') || ($function === null || $function === '')) {
            if (!($column->TotalsRowLabel === null || $column->TotalsRowLabel === '')) {
                $element->SetAttributeValue('totalsRowLabel', $column->TotalsRowLabel);
            }
        }

        if ($function === 'custom' && !($column->TotalsRowFormula === null || $column->TotalsRowFormula === '')) {
            $element->SetAttributeValue('totalsRowFormula', $column->TotalsRowFormula);
        }

        return $element;
    }

    private static function BuildTableStyleInfoElement($model)
    {
        if ($model->TableStyleName === null || $model->TableStyleName === '') {
            return null;
        }

        $element = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'tableStyleInfo',
            new XAttribute('name', $model->TableStyleName),
            new XAttribute('showFirstColumn', $model->ShowFirstColumn ? '1' : '0'),
            new XAttribute('showLastColumn', $model->ShowLastColumn ? '1' : '0'),
            new XAttribute('showRowStripes', $model->ShowRowStripes ? '1' : '0'),
            new XAttribute('showColumnStripes', $model->ShowColumnStripes ? '1' : '0')
        );

        return $element;
    }

    public static function BuildTablePartsElement(WorksheetModel $worksheet, $tableRIdStart)
    {
        if (count($worksheet->ListObjects) === 0) {
            return null;
        }

        $tablePartsElement = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'tableParts',
            new XAttribute('count', count($worksheet->ListObjects))
        );

        for ($i = 0; $i < count($worksheet->ListObjects); $i++) {
            $rId = 'rId' . (string)($tableRIdStart + $i);
            $tablePartsElement->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'tablePart', new XAttribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id', $rId)));
        }

        return $tablePartsElement;
    }

    public static function BuildWorksheetRelationshipsDocument(WorksheetModel $worksheet, $tableFileOffset, $drawingNumber, $commentFileNumber)
    {
        $relationships = new XElement(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationships');
        $rId = 1;

        $orderedHyperlinks = XlsxWorkbookHyperlinks::GetOrderedHyperlinks($worksheet->Hyperlinks);
        for ($i = 0; $i < count($orderedHyperlinks); $i++) {
            $hyperlink = $orderedHyperlinks[$i];
            if ($hyperlink->Address === null || $hyperlink->Address === '') {
                continue;
            }

            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (string)$rId),
                new XAttribute('Type', XlsxWorkbookHyperlinks::HyperlinkRelationshipType),
                new XAttribute('Target', $hyperlink->Address),
                new XAttribute('TargetMode', 'External')
            ));
            $rId++;
        }

        for ($t = 0; $t < count($worksheet->ListObjects); $t++) {
            $globalTableNumber = $tableFileOffset + $t + 1;
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (string)$rId),
                new XAttribute('Type', self::TableRelationshipType),
                new XAttribute('Target', '../tables/table' . (string)$globalTableNumber . '.xml')
            ));
            $rId++;
        }

        if ($drawingNumber > 0) {
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (string)$rId),
                new XAttribute('Type', XlsxWorkbookPictures::DrawingRelationshipType),
                new XAttribute('Target', '../drawings/drawing' . (string)$drawingNumber . '.xml')
            ));
            $rId++;
        }

        if ($commentFileNumber > 0) {
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (string)$rId),
                new XAttribute('Type', XlsxWorkbookComments::CommentsRelationshipType),
                new XAttribute('Target', '../comments' . (string)$commentFileNumber . '.xml')
            ));
            $rId++;
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (string)$rId),
                new XAttribute('Type', XlsxWorkbookComments::VmlDrawingRelationshipType),
                new XAttribute('Target', '../drawings/vmlDrawing' . (string)$commentFileNumber . '.vml')
            ));
            $rId++;
        }

        if (!$relationships->HasElements) {
            return null;
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $relationships);
    }

    public static function LoadWorksheetTableTargets(ZipArchive $archive, $worksheetUri)
    {
        $relsUri = self::GetWorksheetRelsUri($worksheetUri);
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $relsUri);
        if ($entry === null) {
            return array();
        }

        $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        $targets = array();
        $elements = $document->Root === null ? array() : $document->Root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship');
        foreach ($elements as $rel) {
            $id = (string)$rel->Attribute('Id');
            $type = (string)$rel->Attribute('Type');
            $target = (string)$rel->Attribute('Target');
            if ($id === '' || $target === '') {
                continue;
            }

            if (strcasecmp($type, self::TableRelationshipType) !== 0) {
                continue;
            }

            $targets[$id] = XlsxWorkbookArchiveHelpers::ResolvePartUri($worksheetUri, $target);
        }

        return $targets;
    }

    public static function LoadTables(WorksheetModel $worksheetModel, XElement $worksheetRoot, ZipArchive $archive, $worksheetUri, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->ListObjects = array();

        $tablePartsElement = $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'tableParts');
        if ($tablePartsElement === null) {
            return;
        }

        $tableTargets = self::LoadWorksheetTableTargets($archive, $worksheetUri);

        $tablePartElements = $tablePartsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'tablePart');
        foreach ($tablePartElements as $tablePartElement) {
            $rId = (string)$tablePartElement->Attribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id');
            if ($rId === null || $rId === '') {
                continue;
            }

            if (!array_key_exists($rId, $tableTargets)) {
                $issue = new LoadIssue('TBL-L002', DiagnosticSeverity::LossyRecoverable, "Table relationship '" . $rId . "' could not be resolved and the table was dropped.", true, true);
                $issue->SheetName = $sheetName;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $tableUri = $tableTargets[$rId];
            $tableEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $tableUri);
            if ($tableEntry === null) {
                $issue = new LoadIssue('TBL-R001', DiagnosticSeverity::Recoverable, "Table part '" . $tableUri . "' was not found and the table was skipped.", true);
                $issue->SheetName = $sheetName;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $tableDocument = XlsxWorkbookArchiveHelpers::LoadDocument($tableEntry);
            $tableModel = self::LoadTableModel($tableDocument, $diagnostics, $options, $sheetName, $worksheetModel);
            if ($tableModel !== null) {
                $worksheetModel->ListObjects[] = $tableModel;
            }
        }
    }

    private static function LoadTableModel(XDocument $tableDocument, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, WorksheetModel $worksheetModel)
    {
        $root = $tableDocument->Root;
        if ($root === null) {
            return null;
        }

        $tableRef = (string)$root->Attribute('ref');
        if ($tableRef === null || $tableRef === '') {
            $issue = new LoadIssue('TBL-L001', DiagnosticSeverity::LossyRecoverable, 'Table ref attribute is missing and the table was dropped.', true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        $startRow = 0;
        $startColumn = 0;
        $endRow = 0;
        $endColumn = 0;
        if (!self::TryParseTableRef($tableRef, $startRow, $startColumn, $endRow, $endColumn)) {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException("Table ref '" . $tableRef . "' is invalid.");
            }

            $issue = new LoadIssue('TBL-L001', DiagnosticSeverity::LossyRecoverable, "Table ref '" . $tableRef . "' is invalid and the table was dropped.", true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        $model = new ListObjectModel();
        $model->StartRow = $startRow;
        $model->StartColumn = $startColumn;
        $model->EndRow = $endRow;
        $model->EndColumn = $endColumn;

        $displayNameAttribute = $root->Attribute('displayName');
        $model->DisplayName = $displayNameAttribute === null ? '' : (string)$displayNameAttribute;
        $nameAttribute = $root->Attribute('name');
        $model->Name = $nameAttribute === null ? $model->DisplayName : (string)$nameAttribute;
        $commentAttribute = $root->Attribute('comment');
        $model->Comment = $commentAttribute === null ? '' : (string)$commentAttribute;

        if ($model->DisplayName === null || $model->DisplayName === '') {
            $model->DisplayName = $model->Name;
        }

        $headerRowCount = XlsxWorkbookArchiveHelpers::ParseIntAttribute($root->Attribute('headerRowCount'));
        $model->ShowHeaderRow = $headerRowCount === null || $headerRowCount !== 0;

        $totalsRowCount = XlsxWorkbookArchiveHelpers::ParseIntAttribute($root->Attribute('totalsRowCount'));
        $model->ShowTotals = $totalsRowCount !== null && $totalsRowCount > 0;

        $autoFilterElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'autoFilter');
        $model->HasAutoFilter = $autoFilterElement !== null;

        self::LoadTableColumns($model, $root, $diagnostics, $options, $sheetName);
        self::LoadTableStyleInfo($model, $root);

        return $model;
    }

    private static function LoadTableColumns(ListObjectModel $model, XElement $root, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $tableColumnsElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'tableColumns');
        if ($tableColumnsElement === null) {
            $expectedCount = $model->EndColumn - $model->StartColumn + 1;
            for ($c = 0; $c < $expectedCount; $c++) {
                $model->Columns[] = new ListColumnModel($c + 1, 'Column' . (string)($c + 1));
            }

            $issue = new LoadIssue('TBL-R004', DiagnosticSeverity::Recoverable, 'Table columns element was missing; default column names were synthesized.', true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        $columnElements = $tableColumnsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'tableColumn');
        foreach ($columnElements as $columnElement) {
            $idAttr = XlsxWorkbookArchiveHelpers::ParseIntAttribute($columnElement->Attribute('id'));
            $nameAttribute = $columnElement->Attribute('name');
            $name = $nameAttribute === null ? '' : (string)$nameAttribute;
            $columnId = $idAttr !== null && $idAttr > 0 ? $idAttr : count($model->Columns) + 1;

            $columnModel = new ListColumnModel($columnId, $name);
            $totalsRowFunctionAttribute = $columnElement->Attribute('totalsRowFunction');
            $columnModel->TotalsRowFunction = $totalsRowFunctionAttribute === null ? 'none' : (string)$totalsRowFunctionAttribute;
            $totalsRowLabelAttribute = $columnElement->Attribute('totalsRowLabel');
            $columnModel->TotalsRowLabel = $totalsRowLabelAttribute === null ? '' : (string)$totalsRowLabelAttribute;
            $totalsRowFormulaAttribute = $columnElement->Attribute('totalsRowFormula');
            $columnModel->TotalsRowFormula = $totalsRowFormulaAttribute === null ? '' : (string)$totalsRowFormulaAttribute;

            $model->Columns[] = $columnModel;
        }
    }

    private static function LoadTableStyleInfo(ListObjectModel $model, XElement $root)
    {
        $styleInfoElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'tableStyleInfo');
        if ($styleInfoElement === null) {
            return;
        }

        $nameAttribute = $styleInfoElement->Attribute('name');
        $model->TableStyleName = $nameAttribute === null ? '' : (string)$nameAttribute;
        $model->ShowFirstColumn = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($styleInfoElement->Attribute('showFirstColumn'));
        $model->ShowLastColumn = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($styleInfoElement->Attribute('showLastColumn'));
        $model->ShowRowStripes = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($styleInfoElement->Attribute('showRowStripes'));
        $model->ShowColumnStripes = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($styleInfoElement->Attribute('showColumnStripes'));
    }

    private static function TryParseTableRef($tableRef, &$startRow, &$startColumn, &$endRow, &$endColumn)
    {
        $startRow = 0;
        $startColumn = 0;
        $endRow = 0;
        $endColumn = 0;

        if ($tableRef === null || trim($tableRef) === '') {
            return false;
        }

        $parts = explode(':', $tableRef);
        if (count($parts) !== 2) {
            return false;
        }

        $first = null;
        $last = null;
        if (!XlsxWorkbookSerializerCommon::TryParseCellReference($parts[0], $first) || !XlsxWorkbookSerializerCommon::TryParseCellReference($parts[1], $last)) {
            return false;
        }

        if ($last->RowIndex < $first->RowIndex || $last->ColumnIndex < $first->ColumnIndex) {
            return false;
        }

        $startRow = $first->RowIndex;
        $startColumn = $first->ColumnIndex;
        $endRow = $last->RowIndex;
        $endColumn = $last->ColumnIndex;
        return true;
    }

    private static function GetWorksheetRelsUri($worksheetUri)
    {
        $normalized = ltrim($worksheetUri, '/');
        $slashIndex = strrpos($normalized, '/');
        $directory = $slashIndex === false ? '' : substr($normalized, 0, $slashIndex + 1);
        $fileName = $slashIndex === false ? $normalized : substr($normalized, $slashIndex + 1);
        return '/' . $directory . '_rels/' . $fileName . '.rels';
    }
}
