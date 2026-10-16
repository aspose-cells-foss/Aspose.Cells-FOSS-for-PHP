<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\HyperlinkModel;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookHyperlinks
{
    const HyperlinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

    public static function BuildHyperlinksElement(WorksheetModel $worksheet)
    {
        $orderedHyperlinks = self::GetOrderedHyperlinks($worksheet->Hyperlinks);
        if (count($orderedHyperlinks) === 0) {
            return null;
        }

        $relationshipId = 1;
        $hyperlinksElement = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'hyperlinks');
        for ($index = 0; $index < count($orderedHyperlinks); $index++) {
            $hyperlink = $orderedHyperlinks[$index];
            if (!self::ShouldPersistHyperlink($hyperlink)) {
                continue;
            }

            $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'hyperlink', new XAttribute('ref', self::ToRangeReference($hyperlink)));
            if (!($hyperlink->Address === null || $hyperlink->Address === '')) {
                $element->SetAttributeValue(XlsxWorkbookSerializerCommon::RelationshipNs . 'id', 'rId' . (string)$relationshipId);
                $relationshipId++;
            }

            if (!($hyperlink->SubAddress === null || $hyperlink->SubAddress === '')) {
                $element->SetAttributeValue('location', $hyperlink->SubAddress);
            }

            if (!($hyperlink->ScreenTip === null || $hyperlink->ScreenTip === '')) {
                $element->SetAttributeValue('tooltip', $hyperlink->ScreenTip);
            }

            if (!($hyperlink->TextToDisplay === null || $hyperlink->TextToDisplay === '')) {
                $element->SetAttributeValue('display', $hyperlink->TextToDisplay);
            }

            $hyperlinksElement->Add($element);
        }

        if (!$hyperlinksElement->HasElements) {
            return null;
        }

        return $hyperlinksElement;
    }

    public static function BuildWorksheetHyperlinkRelationships(WorksheetModel $worksheet)
    {
        $orderedHyperlinks = self::GetOrderedHyperlinks($worksheet->Hyperlinks);
        $relationships = new XElement(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationships');
        $relationshipId = 1;

        for ($index = 0; $index < count($orderedHyperlinks); $index++) {
            $hyperlink = $orderedHyperlinks[$index];
            if ($hyperlink->Address === null || $hyperlink->Address === '') {
                continue;
            }

            $relationships->Add(new XElement(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (string)$relationshipId),
                new XAttribute('Type', self::HyperlinkRelationshipType),
                new XAttribute('Target', $hyperlink->Address),
                new XAttribute('TargetMode', 'External')));
            $relationshipId++;
        }

        if (!$relationships->HasElements) {
            return null;
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $relationships);
    }

    public static function LoadWorksheetHyperlinkTargets(ZipArchive $archive, $worksheetUri)
    {
        $relationshipsEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, self::GetWorksheetRelationshipsUri($worksheetUri));
        if ($relationshipsEntry === null) {
            return array();
        }

        $document = XlsxWorkbookArchiveHelpers::LoadDocument($relationshipsEntry);
        $targets = array();
        $root = $document->Root;
        $elements = $root === null ? array() : $root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship');
        foreach ($elements as $relationship) {
            $id = (string)$relationship->Attribute('Id');
            $type = (string)$relationship->Attribute('Type');
            $target = (string)$relationship->Attribute('Target');
            if ($id === '' || $target === '' || strcasecmp($type, self::HyperlinkRelationshipType) !== 0) {
                continue;
            }

            $targets[$id] = $target;
        }

        return $targets;
    }

    public static function LoadHyperlinks(WorksheetModel $worksheetModel, XElement $worksheetRoot, $hyperlinkTargets, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->Hyperlinks = array();

        $hyperlinksElement = $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'hyperlinks');
        $elements = $hyperlinksElement === null ? array() : $hyperlinksElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'hyperlink');
        foreach ($elements as $hyperlinkElement) {
            $reference = (string)$hyperlinkElement->Attribute('ref');
            $candidate = null;
            if (!self::TryParseHyperlinkReference($reference, $candidate)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The hyperlink reference '" . $reference . "' is invalid.");
                }

                $issue = new LoadIssue('HL-L001', DiagnosticSeverity::LossyRecoverable, 'A hyperlink with an invalid reference was dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $reference;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            if (self::ContainsOverlappingHyperlink($worksheetModel->Hyperlinks, $candidate)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The hyperlink reference '" . $reference . "' overlaps an existing hyperlink.");
                }

                $issue = new LoadIssue('HL-L003', DiagnosticSeverity::LossyRecoverable, 'Overlapping hyperlink ranges were normalized during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $reference;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $candidate->SubAddress = self::NormalizeAttributeValue((string)$hyperlinkElement->Attribute('location'));
            $candidate->ScreenTip = self::NormalizeAttributeValue((string)$hyperlinkElement->Attribute('tooltip'));
            $candidate->TextToDisplay = self::NormalizeAttributeValue((string)$hyperlinkElement->Attribute('display'));

            $relationshipId = (string)$hyperlinkElement->Attribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id');
            if (!($relationshipId === null || $relationshipId === '')) {
                if (array_key_exists($relationshipId, $hyperlinkTargets)) {
                    $candidate->Address = self::NormalizeAttributeValue($hyperlinkTargets[$relationshipId]);
                } elseif ($options->StrictMode) {
                    throw new InvalidFileFormatException("The hyperlink relationship '" . $relationshipId . "' could not be resolved.");
                } else {
                    $issue = new LoadIssue('HL-L002', DiagnosticSeverity::LossyRecoverable, 'A hyperlink relationship was missing and the external address was cleared.', true, true);
                    $issue->SheetName = $sheetName;
                    $issue->CellRef = $reference;
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                }
            }

            if (($candidate->Address === null || $candidate->Address === '') && ($candidate->SubAddress === null || $candidate->SubAddress === '')) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The hyperlink reference '" . $reference . "' does not define an address or sub-address.");
                }

                $issue = new LoadIssue('HL-L004', DiagnosticSeverity::LossyRecoverable, 'A hyperlink without an address or sub-address was dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $reference;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $worksheetModel->Hyperlinks[] = $candidate;
        }
    }

    public static function GetOrderedHyperlinks($hyperlinks)
    {
        $ordered = array();
        foreach ($hyperlinks as $hyperlink) {
            $ordered[] = $hyperlink;
        }

        usort($ordered, array(__CLASS__, 'CompareHyperlinks'));
        return $ordered;
    }

    private static function CompareHyperlinks($left, $right)
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

    private static function ShouldPersistHyperlink($hyperlink)
    {
        return !($hyperlink->Address === null || $hyperlink->Address === '') || !($hyperlink->SubAddress === null || $hyperlink->SubAddress === '');
    }

    private static function ToRangeReference($hyperlink)
    {
        $first = (new CellAddress($hyperlink->FirstRow, $hyperlink->FirstColumn))->ToString();
        if ($hyperlink->TotalRows === 1 && $hyperlink->TotalColumns === 1) {
            return $first;
        }

        $last = (new CellAddress($hyperlink->FirstRow + $hyperlink->TotalRows - 1, $hyperlink->FirstColumn + $hyperlink->TotalColumns - 1))->ToString();
        return $first . ':' . $last;
    }

    private static function GetWorksheetRelationshipsUri($worksheetUri)
    {
        $normalizedUri = ltrim($worksheetUri, '/');
        $slashIndex = strrpos($normalizedUri, '/');
        $directory = $slashIndex === false ? '' : substr($normalizedUri, 0, $slashIndex + 1);
        $fileName = $slashIndex === false ? $normalizedUri : substr($normalizedUri, $slashIndex + 1);
        return '/' . $directory . '_rels/' . $fileName . '.rels';
    }

    private static function TryParseHyperlinkReference($reference, &$hyperlink)
    {
        $hyperlink = new HyperlinkModel();
        $region = null;
        if (!XlsxWorkbookSerializerCommon::TryParseMergeReference($reference === null ? '' : $reference, $region)) {
            return false;
        }

        $hyperlink->FirstRow = $region->FirstRow;
        $hyperlink->FirstColumn = $region->FirstColumn;
        $hyperlink->TotalRows = $region->TotalRows;
        $hyperlink->TotalColumns = $region->TotalColumns;
        return true;
    }

    private static function ContainsOverlappingHyperlink($hyperlinks, $candidate)
    {
        for ($index = 0; $index < count($hyperlinks); $index++) {
            if (self::HyperlinksOverlap($hyperlinks[$index], $candidate)) {
                return true;
            }
        }

        return false;
    }

    private static function HyperlinksOverlap($left, $right)
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

    private static function NormalizeAttributeValue($value)
    {
        return ($value === null || $value === '') ? null : $value;
    }
}
