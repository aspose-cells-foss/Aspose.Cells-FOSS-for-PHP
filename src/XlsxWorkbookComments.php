<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CommentModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

// Provides comment load and save helpers for the XLSX workbook writer.
final class XlsxWorkbookComments
{
    const CommentsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
    const VmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
    const CommentsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
    const VmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing';

    private function __construct()
    {
    }

    public static function BuildCommentsDocument(WorksheetModel $worksheet)
    {
        $sorted = self::SortComments($worksheet->Comments);

        $authors = array();
        $authorToIndex = array();

        for ($i = 0; $i < count($sorted); $i++) {
            $author = $sorted[$i]->Author;
            if ($author === null) {
                $author = '';
            }

            if (!array_key_exists($author, $authorToIndex)) {
                $authorToIndex[$author] = count($authors);
                $authors[] = $author;
            }
        }

        $authorsElement = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'authors');
        for ($a = 0; $a < count($authors); $a++) {
            $authorsElement->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'author', $authors[$a]));
        }

        $commentListElement = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'commentList');
        for ($i = 0; $i < count($sorted); $i++) {
            $comment = $sorted[$i];
            $author = $comment->Author === null ? '' : $comment->Author;
            $authorId = 0;
            if (array_key_exists($author, $authorToIndex)) {
                $authorId = $authorToIndex[$author];
            }

            $cellRef = (new CellAddress($comment->Row, $comment->Column))->ToString();
            $note = $comment->Note === null ? '' : $comment->Note;

            $tElement = new XElement(XlsxWorkbookSerializerCommon::MainNs . 't', $note);
            if (XlsxWorkbookSerializerCommon::NeedsPreserveWhitespace($note)) {
                $tElement->SetAttributeValue(XlsxWorkbookSerializerCommon::XmlNs . 'space', 'preserve');
            }

            $commentElement = new XElement(
                XlsxWorkbookSerializerCommon::MainNs . 'comment',
                new XAttribute('ref', $cellRef),
                new XAttribute('authorId', (string) $authorId),
                new XElement(
                    XlsxWorkbookSerializerCommon::MainNs . 'text',
                    new XElement(XlsxWorkbookSerializerCommon::MainNs . 'r', $tElement)
                )
            );

            $commentListElement->Add($commentElement);
        }

        $root = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'comments', $authorsElement, $commentListElement);
        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $root);
    }

    public static function WriteVmlDrawing(ZipArchive $archive, WorksheetModel $worksheet, $commentFileNumber)
    {
        $path = 'xl/drawings/vmlDrawing' . $commentFileNumber . '.vml';
        $archive->addEntry($path, self::BuildVmlContent($worksheet));
    }

    private static function BuildVmlContent(WorksheetModel $worksheet)
    {
        $sb = '';
        $sb .= '<xml xmlns:v="urn:schemas-microsoft-com:vml"';
        $sb .= ' xmlns:o="urn:schemas-microsoft-com:office:office"';
        $sb .= ' xmlns:x="urn:schemas-microsoft-com:office:excel">';
        $sb .= '<o:shapelayout v:ext="edit">';
        $sb .= '<o:idmap v:ext="edit" data="1"/>';
        $sb .= '</o:shapelayout>';
        $sb .= '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"';
        $sb .= ' path="m,l,21600r21600,l21600,xe">';
        $sb .= '<v:stroke joinstyle="miter"/>';
        $sb .= '<v:path gradientshapeok="t" o:connecttype="rect"/>';
        $sb .= '</v:shapetype>';

        $sorted = self::SortComments($worksheet->Comments);
        for ($i = 0; $i < count($sorted); $i++) {
            $comment = $sorted[$i];
            $shapeId = 1025 + $i;
            if ($comment->RawVmlShapeXml !== null) {
                $sb .= self::PatchVmlShape($comment->RawVmlShapeXml, $comment, $shapeId);
            } else {
                $sb .= self::BuildDefaultVmlShape($comment, $shapeId);
            }
        }

        $sb .= '</xml>';
        return $sb;
    }

    private static function PatchVmlShape($raw, CommentModel $comment, $shapeId)
    {
        $raw = self::ReplaceInlineValue($raw, 'id="_x0000_s', '"', (string) $shapeId);

        $widthPt = self::FormatPt($comment->Width);
        $heightPt = self::FormatPt($comment->Height);
        $raw = self::ReplaceInlineValue($raw, 'width:', 'pt', $widthPt);
        $raw = self::ReplaceInlineValue($raw, 'height:', 'pt', $heightPt);

        $raw = self::SetStyleVisibility($raw, $comment->IsVisible);
        $raw = self::ReplaceTagContent($raw, 'x:Row', (string) $comment->Row);
        $raw = self::ReplaceTagContent($raw, 'x:Column', (string) $comment->Column);
        $raw = self::SetClientDataVisible($raw, $comment->IsVisible);

        return $raw;
    }

    private static function BuildDefaultVmlShape(CommentModel $comment, $shapeId)
    {
        $widthPt = self::FormatPt($comment->Width);
        $heightPt = self::FormatPt($comment->Height);
        $visStyle = $comment->IsVisible ? ';visibility:visible' : '';

        $anchorLeft = $comment->Column + 1;
        $anchorTop = $comment->Row > 0 ? $comment->Row - 1 : 0;
        $anchorRight = $comment->Column + 4;
        $anchorBottom = $comment->Row + 5;

        $sb = '';
        $sb .= '<v:shape id="_x0000_s';
        $sb .= (string) $shapeId;
        $sb .= '" type="#_x0000_t202"';
        $sb .= ' style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:';
        $sb .= $widthPt;
        $sb .= 'pt;height:';
        $sb .= $heightPt;
        $sb .= 'pt;z-index:1';
        $sb .= $visStyle;
        $sb .= '" fillcolor="#ffffe1" o:insetmode="auto">';
        $sb .= '<v:fill color2="#ffffe1"/>';
        $sb .= '<v:shadow on="t" color="black" obscured="t"/>';
        $sb .= '<v:path o:connecttype="none"/>';
        $sb .= '<v:textbox style="mso-direction-alt:auto">';
        $sb .= '<div style="text-align:left"/>';
        $sb .= '</v:textbox>';
        $sb .= '<x:ClientData ObjectType="Note">';
        $sb .= '<x:MoveWithCells/>';
        $sb .= '<x:SizeWithCells/>';
        $sb .= '<x:Anchor>';
        $sb .= (string) $anchorLeft;
        $sb .= ',15,';
        $sb .= (string) $anchorTop;
        $sb .= ',2,';
        $sb .= (string) $anchorRight;
        $sb .= ',15,';
        $sb .= (string) $anchorBottom;
        $sb .= ',16</x:Anchor>';
        $sb .= '<x:AutoFill>False</x:AutoFill>';
        $sb .= '<x:Row>';
        $sb .= (string) $comment->Row;
        $sb .= '</x:Row>';
        $sb .= '<x:Column>';
        $sb .= (string) $comment->Column;
        $sb .= '</x:Column>';
        if ($comment->IsVisible) {
            $sb .= '<x:Visible/>';
        }

        $sb .= '</x:ClientData>';
        $sb .= '</v:shape>';
        return $sb;
    }

    public static function FormatPt($pixels)
    {
        $pt = $pixels * 0.75;
        $formatted = number_format($pt, 2, '.', '');
        $formatted = rtrim(rtrim($formatted, '0'), '.');
        return $formatted;
    }

    public static function SortComments($comments)
    {
        $sorted = $comments;
        for ($i = 0; $i < count($sorted) - 1; $i++) {
            for ($j = $i + 1; $j < count($sorted); $j++) {
                $a = $sorted[$i];
                $b = $sorted[$j];
                if ($a->Row > $b->Row || ($a->Row === $b->Row && $a->Column > $b->Column)) {
                    $sorted[$i] = $b;
                    $sorted[$j] = $a;
                }
            }
        }

        return $sorted;
    }

    // Replaces the value between prefix and the next occurrence of suffix.
    public static function ReplaceInlineValue($text, $prefix, $suffix, $newValue)
    {
        $start = strpos($text, $prefix);
        if ($start === false) {
            return $text;
        }

        $valueStart = $start + strlen($prefix);
        $end = strpos($text, $suffix, $valueStart);
        if ($end === false) {
            return $text;
        }

        return substr($text, 0, $valueStart) . $newValue . substr($text, $end);
    }

    // Replaces the text content of <tagName>...</tagName> (first occurrence).
    public static function ReplaceTagContent($text, $tagName, $newContent)
    {
        $openTag = '<' . $tagName . '>';
        $closeTag = '</' . $tagName . '>';
        $start = strpos($text, $openTag);
        if ($start === false) {
            return $text;
        }

        $contentStart = $start + strlen($openTag);
        $end = strpos($text, $closeTag, $contentStart);
        if ($end === false) {
            return $text;
        }

        return substr($text, 0, $contentStart) . $newContent . substr($text, $end);
    }

    private static function SetStyleVisibility($raw, $isVisible)
    {
        $hasVisible = strpos($raw, 'visibility:visible') !== false;

        if ($isVisible && !$hasVisible) {
            $stylePrefix = 'style="';
            $styleStart = strpos($raw, $stylePrefix);
            if ($styleStart !== false) {
                $quoteStart = $styleStart + strlen($stylePrefix);
                $quoteEnd = strpos($raw, '"', $quoteStart);
                if ($quoteEnd !== false) {
                    return substr($raw, 0, $quoteEnd) . ';visibility:visible' . substr($raw, $quoteEnd);
                }
            }

            return $raw;
        }

        if (!$isVisible && $hasVisible) {
            $withSemi = ';visibility:visible';
            $idx = strpos($raw, $withSemi);
            if ($idx !== false) {
                return substr($raw, 0, $idx) . substr($raw, $idx + strlen($withSemi));
            }

            $noSemi = 'visibility:visible';
            $idx = strpos($raw, $noSemi);
            if ($idx !== false) {
                return substr($raw, 0, $idx) . substr($raw, $idx + strlen($noSemi));
            }
        }

        return $raw;
    }

    private static function SetClientDataVisible($raw, $isVisible)
    {
        $hasTag = strpos($raw, '<x:Visible/>') !== false;

        if ($isVisible && !$hasTag) {
            $endTag = '</x:ClientData>';
            $idx = strpos($raw, $endTag);
            if ($idx !== false) {
                return substr($raw, 0, $idx) . '<x:Visible/>' . substr($raw, $idx);
            }
        } elseif (!$isVisible && $hasTag) {
            return str_replace('<x:Visible/>', '', $raw);
        }

        return $raw;
    }
}
