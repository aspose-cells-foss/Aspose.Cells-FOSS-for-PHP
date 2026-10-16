<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CommentModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Io\ZipEntry;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

// Loads worksheet comments (legacy notes) from the package.
final class XlsxWorkbookCommentsLoader
{
    private function __construct()
    {
    }

    public static function LoadComments(WorksheetModel $worksheetModel, ZipArchive $archive, $worksheetUri, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->Comments = array();

        $relsUri = self::GetWorksheetRelsUri($worksheetUri);
        $relsEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $relsUri);
        if ($relsEntry === null) {
            return;
        }

        $relsDoc = null;
        try {
            $relsDoc = XlsxWorkbookArchiveHelpers::LoadDocument($relsEntry);
        } catch (\Throwable $exception) {
            return;
        }

        if ($relsDoc->Root === null) {
            return;
        }

        $commentsTarget = null;
        $vmlTarget = null;

        foreach ($relsDoc->Root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship') as $rel) {
            $type = (string) $rel->Attribute('Type');
            $target = (string) $rel->Attribute('Target');
            if ($type === null || $type === '' || $target === null || $target === '') {
                continue;
            }

            if (strcasecmp($type, XlsxWorkbookComments::CommentsRelationshipType) === 0) {
                $commentsTarget = XlsxWorkbookArchiveHelpers::ResolvePartUri($worksheetUri, $target);
            } elseif (strcasecmp($type, XlsxWorkbookComments::VmlDrawingRelationshipType) === 0) {
                $vmlTarget = XlsxWorkbookArchiveHelpers::ResolvePartUri($worksheetUri, $target);
            }
        }

        if ($commentsTarget === null) {
            if ($vmlTarget !== null) {
                $issue = new LoadIssue('COM-R003', DiagnosticSeverity::Recoverable, 'A VML drawing part was found without a comments relationship; comments were skipped.', false);
                $issue->SheetName = $sheetName;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            }

            return;
        }

        $commentsEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $commentsTarget);
        if ($commentsEntry === null) {
            $issue = new LoadIssue('COM-R002', DiagnosticSeverity::Recoverable, "Comments part '" . $commentsTarget . "' was referenced but not found.", false);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        self::LoadCommentsXml($worksheetModel, $commentsEntry, $diagnostics, $options, $sheetName);

        if ($vmlTarget === null) {
            $issue = new LoadIssue('COM-R002', DiagnosticSeverity::Recoverable, 'Comments were loaded but the VML drawing part is missing; default shape dimensions will be used.', true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        $vmlEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $vmlTarget);
        if ($vmlEntry === null) {
            $issue = new LoadIssue('COM-R002', DiagnosticSeverity::Recoverable, "VML drawing part '" . $vmlTarget . "' was referenced but not found; default shape dimensions will be used.", true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        self::LoadVmlDrawing($worksheetModel, $vmlEntry);
    }

    private static function LoadCommentsXml(WorksheetModel $worksheetModel, ZipEntry $entry, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $document = null;
        try {
            $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        } catch (\Throwable $exception) {
            return;
        }

        if ($document->Root === null) {
            return;
        }

        $authors = array();
        $authorsElement = $document->Root->Element(XlsxWorkbookSerializerCommon::MainNs . 'authors');
        if ($authorsElement !== null) {
            foreach ($authorsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'author') as $authorElement) {
                $authorValue = $authorElement->Value;
                $authors[] = $authorValue === null ? '' : $authorValue;
            }
        }

        $commentListElement = $document->Root->Element(XlsxWorkbookSerializerCommon::MainNs . 'commentList');
        if ($commentListElement === null) {
            return;
        }

        foreach ($commentListElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'comment') as $commentElement) {
            $cellRef = (string) $commentElement->Attribute('ref');
            if ($cellRef === null || trim($cellRef) === '') {
                $issue = new LoadIssue('COM-R001', DiagnosticSeverity::Recoverable, 'A comment with a missing ref attribute was skipped.', false);
                $issue->SheetName = $sheetName;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $address = null;
            if (!XlsxWorkbookSerializerCommon::TryParseCellReference($cellRef, $address)) {
                $issue = new LoadIssue('COM-R001', DiagnosticSeverity::Recoverable, "A comment with invalid ref '" . $cellRef . "' was skipped.", false);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $cellRef;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $author = '';
            $authorIdStr = (string) $commentElement->Attribute('authorId');
            $authorId = 0;
            if (!($authorIdStr === null || $authorIdStr === '')
                && is_numeric($authorIdStr)
                && (int) $authorIdStr >= 0
                && (int) $authorIdStr < count($authors)) {
                $authorId = (int) $authorIdStr;
                $author = $authors[$authorId];
            }

            $textElement = $commentElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'text');
            $note = self::ReadNoteText($textElement);

            $model = new CommentModel();
            $model->Row = $address->RowIndex;
            $model->Column = $address->ColumnIndex;
            $model->Author = $author;
            $model->Note = $note;
            $worksheetModel->Comments[] = $model;
        }
    }

    private static function ReadNoteText($textElement)
    {
        if ($textElement === null) {
            return '';
        }

        $hasRuns = false;
        $sb = '';
        foreach ($textElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'r') as $runElement) {
            $hasRuns = true;
            $tElement = $runElement->Element(XlsxWorkbookSerializerCommon::MainNs . 't');
            if ($tElement !== null) {
                $sb .= $tElement->Value;
            }
        }

        if (!$hasRuns) {
            return $textElement->Value;
        }

        return $sb;
    }

    private static function LoadVmlDrawing(WorksheetModel $worksheetModel, ZipEntry $entry)
    {
        $vmlText = null;
        try {
            $vmlText = $entry->getContent();
        } catch (\Throwable $exception) {
            return;
        }

        $commentMap = array();
        for ($i = 0; $i < count($worksheetModel->Comments); $i++) {
            $c = $worksheetModel->Comments[$i];
            $key = ($c->Row << 20) | $c->Column;
            if (!array_key_exists($key, $commentMap)) {
                $commentMap[$key] = $c;
            }
        }

        $shapes = self::ExtractVmlShapes($vmlText);
        for ($s = 0; $s < count($shapes); $s++) {
            $shapeText = $shapes[$s];

            if (stripos($shapeText, 'ObjectType="Note"') === false) {
                continue;
            }

            $row = self::ExtractTagInt($shapeText, 'x:Row');
            $col = self::ExtractTagInt($shapeText, 'x:Column');
            if ($row < 0 || $col < 0) {
                continue;
            }

            $mapKey = ($row << 20) | $col;
            if (!array_key_exists($mapKey, $commentMap)) {
                continue;
            }

            $model = $commentMap[$mapKey];

            $styleAttr = self::ExtractAttributeValue($shapeText, 'style="');
            if ($styleAttr !== null) {
                $model->IsVisible = stripos($styleAttr, 'visibility:visible') !== false;

                $w = self::ParseStylePt($styleAttr, 'width:');
                if ($w > 0) {
                    $model->Width = $w;
                }

                $h = self::ParseStylePt($styleAttr, 'height:');
                if ($h > 0) {
                    $model->Height = $h;
                }
            }

            $model->RawVmlShapeXml = $shapeText;
        }
    }

    private static function ExtractVmlShapes($vmlText)
    {
        $shapes = array();
        $searchStart = 0;
        $closeTag = '</v:shape>';

        while (true) {
            $shapeStart = self::FindNextVmlShape($vmlText, $searchStart);
            if ($shapeStart < 0) {
                break;
            }

            $shapeEnd = strpos($vmlText, $closeTag, $shapeStart);
            if ($shapeEnd === false) {
                break;
            }

            $endPos = $shapeEnd + strlen($closeTag);
            $shapes[] = substr($vmlText, $shapeStart, $endPos - $shapeStart);
            $searchStart = $endPos;
        }

        return $shapes;
    }

    // Finds the next <v:shape> start (not <v:shapetype> or any other v:shapeXxx variant).
    private static function FindNextVmlShape($vmlText, $searchFrom)
    {
        $pos = $searchFrom;
        while (true) {
            $idx = strpos($vmlText, '<v:shape', $pos);
            if ($idx === false) {
                return -1;
            }

            $afterTag = $idx + 8; // length of "<v:shape"
            if ($afterTag >= strlen($vmlText)) {
                return -1;
            }

            $next = $vmlText[$afterTag];
            if ($next === ' ' || $next === '>' || $next === "\r" || $next === "\n" || $next === "\t") {
                return $idx;
            }

            $pos = $idx + 1;
        }
    }

    private static function ExtractTagInt($text, $tagName)
    {
        $openTag = '<' . $tagName . '>';
        $closeTag = '</' . $tagName . '>';
        $start = strpos($text, $openTag);
        if ($start === false) {
            return -1;
        }

        $valueStart = $start + strlen($openTag);
        $end = strpos($text, $closeTag, $valueStart);
        if ($end === false) {
            return -1;
        }

        $valuePart = trim(substr($text, $valueStart, $end - $valueStart));
        if (is_numeric($valuePart)) {
            return (int) $valuePart;
        }

        return -1;
    }

    private static function ExtractAttributeValue($text, $prefix)
    {
        $start = strpos($text, $prefix);
        if ($start === false) {
            return null;
        }

        $valueStart = $start + strlen($prefix);
        $end = strpos($text, '"', $valueStart);
        if ($end === false) {
            return null;
        }

        return substr($text, $valueStart, $end - $valueStart);
    }

    private static function ParseStylePt($style, $prefix)
    {
        $start = stripos($style, $prefix);
        if ($start === false) {
            return 0;
        }

        $valueStart = $start + strlen($prefix);
        $end = strpos($style, ';', $valueStart);
        if ($end === false) {
            $end = strpos($style, '"', $valueStart);
        }

        if ($end === false) {
            $end = strlen($style);
        }

        $valuePart = trim(substr($style, $valueStart, $end - $valueStart));
        if (str_ends_with(strtolower($valuePart), 'pt')) {
            $valuePart = substr($valuePart, 0, strlen($valuePart) - 2);
        }

        if (is_numeric($valuePart) && (float) $valuePart > 0) {
            return (int) round((float) $valuePart / 0.75);
        }

        return 0;
    }

    private static function GetWorksheetRelsUri($worksheetUri)
    {
        $normalized = ltrim($worksheetUri, '/');
        $slashIndex = strrpos($normalized, '/');
        $directory = $slashIndex !== false ? substr($normalized, 0, $slashIndex + 1) : '';
        $fileName = $slashIndex !== false ? substr($normalized, $slashIndex + 1) : $normalized;
        return '/' . $directory . '_rels/' . $fileName . '.rels';
    }
}
