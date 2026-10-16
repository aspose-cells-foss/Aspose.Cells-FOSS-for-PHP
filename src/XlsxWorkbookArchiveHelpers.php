<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;

final class XlsxWorkbookArchiveHelpers
{
    public static function EnsureSeekable($stream)
    {
        if (is_string($stream)) {
            return $stream;
        }

        if (is_resource($stream)) {
            $metadata = stream_get_meta_data($stream);
            if ($metadata['seekable']) {
                return $stream;
            }
        } elseif (is_object($stream) && isset($stream->CanSeek) && $stream->CanSeek) {
            return $stream;
        }

        $memory = fopen('php://temp', 'r+b');
        stream_copy_to_stream($stream, $memory);
        rewind($memory);
        return $memory;
    }

    public static function GetEntry(ZipArchive $archive, $uri)
    {
        return $archive->getEntry(ltrim(str_replace('\\', '/', $uri), '/'));
    }

    public static function LoadDocument($entry)
    {
        $rawStream = $entry->Open();
        $stream = self::EnsureSeekable($rawStream);
        fseek($stream, 0);
        return XDocument::Load($stream);
    }

    public static function LoadRelationships(ZipArchive $archive, $relationshipsUri, $sourceUri)
    {
        $entry = self::GetEntry($archive, $relationshipsUri);
        if ($entry === null) {
            return array();
        }

        $document = self::LoadDocument($entry);
        $relationships = array();
        $root = $document->Root;
        if ($root !== null) {
            $elements = $root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship');
            foreach ($elements as $relationship) {
                $id = (string)$relationship->Attribute('Id');
                $target = (string)$relationship->Attribute('Target');
                if ($id === '' || $target === '') {
                    continue;
                }

                $relationships[$id] = self::ResolvePartUri($sourceUri, $target);
            }
        }

        return $relationships;
    }

    public static function FindRelationshipTarget($workbookRelationships, $expectedPath)
    {
        foreach ($workbookRelationships as $value) {
            if (self::RelationshipPointsTo($value, $expectedPath)) {
                return $value;
            }
        }

        return '';
    }

    public static function FindRelationshipTargetByType($workbookRelationships, $relationshipType)
    {
        foreach ($workbookRelationships as $value) {
            if (self::GetRelationshipTargetKind($value) === $relationshipType) {
                return $value;
            }
        }

        return '';
    }

    public static function LoadSharedStrings(ZipArchive $archive, $workbookRelationships, $options, $diagnostics)
    {
        $sharedStringsUri = self::FindRelationshipTarget($workbookRelationships, '/xl/sharedStrings.xml');
        if ($sharedStringsUri === '' || $sharedStringsUri === null) {
            if (count($workbookRelationships) === 0) {
                $sharedStringsUri = '/xl/sharedStrings.xml';
            } else {
                $sharedStringsUri = self::FindRelationshipTargetByType($workbookRelationships, XlsxWorkbookSerializerCommon::SharedStringsRelationshipType);
            }
        }

        if ($sharedStringsUri === '' || $sharedStringsUri === null) {
            $sharedStringsUri = '/xl/sharedStrings.xml';
        }

        $entry = self::GetEntry($archive, $sharedStringsUri);
        if ($entry === null) {
            return array();
        }

        $document = self::LoadDocument($entry);
        $items = array();
        $root = $document->Root;
        if ($root !== null) {
            $elements = $root->Elements(XlsxWorkbookSerializerCommon::MainNs . 'si');
            foreach ($elements as $item) {
                $items[] = self::ReadInlineString($item);
            }
        }

        return $items;
    }

    public static function LoadDateStyleIndexes(ZipArchive $archive, $workbookRelationships, $options, $diagnostics)
    {
        $stylesUri = self::FindRelationshipTarget($workbookRelationships, '/xl/styles.xml');
        if ($stylesUri === '' || $stylesUri === null) {
            $stylesUri = '/xl/styles.xml';
        }

        $entry = self::GetEntry($archive, $stylesUri);
        if ($entry === null) {
            return array();
        }

        $document = self::LoadDocument($entry);
        $customFormats = array();
        $root = $document->Root;
        if ($root !== null) {
            $numFmtsElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'numFmts');
            if ($numFmtsElement !== null) {
                $numFmts = $numFmtsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'numFmt');
                foreach ($numFmts as $numFmt) {
                    $idAttr = $numFmt->Attribute('numFmtId');
                    $codeAttr = $numFmt->Attribute('formatCode');
                    $id = self::ParseIntAttribute($idAttr);
                    $code = $codeAttr !== null ? (string)$codeAttr->Value : null;
                    if ($id !== null && !($code === null || $code === '')) {
                        $customFormats[$id] = $code;
                    }
                }
            }

            $dateStyleIndexes = array();
            $cellXfsElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'cellXfs');
            if ($cellXfsElement !== null) {
                $cellXfs = $cellXfsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'xf');
                $builtIn = XlsxWorkbookSerializerCommon::builtInDateFormats();
                $index = 0;
                foreach ($cellXfs as $xf) {
                    $numFmtId = self::ParseIntAttribute($xf->Attribute('numFmtId'));
                    if ($numFmtId !== null) {
                        if (isset($builtIn[$numFmtId])) {
                            $dateStyleIndexes[$index] = true;
                            $index++;
                            continue;
                        }

                        if (array_key_exists($numFmtId, $customFormats)) {
                            $formatCode = $customFormats[$numFmtId];
                            if (self::LooksLikeDateFormat($formatCode)) {
                                $dateStyleIndexes[$index] = true;
                            }
                        }
                    }
                    $index++;
                }
            }
            return $dateStyleIndexes;
        }

        return array();
    }

    public static function ParseBoolAttribute($attribute)
    {
        if ($attribute === null) {
            return false;
        }

        $value = (string)$attribute->Value;
        return $value === '1' || strcasecmp($value, 'true') === 0;
    }

    public static function ParseIntAttribute($attribute)
    {
        if ($attribute === null) {
            return null;
        }

        $value = trim((string)$attribute->Value);
        if (preg_match('/^[+-]?[0-9]+$/', $value) === 1) {
            return intval($value);
        }

        return null;
    }

    public static function ParseDoubleAttribute($attribute)
    {
        if ($attribute === null) {
            return null;
        }

        $value = trim((string)$attribute->Value);
        if (is_numeric($value)) {
            return floatval($value);
        }

        return null;
    }

    public static function NormalizeFormula($value)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        $normalizedValue = trim($value);
        return str_starts_with($normalizedValue, '=') ? substr($normalizedValue, 1) : $normalizedValue;
    }

    public static function ReadInlineString($inlineStringElement)
    {
        if ($inlineStringElement === null) {
            return '';
        }

        $textNodes = array();
        $elements = $inlineStringElement->DescendantsAndSelf();
        foreach ($elements as $element) {
            if ($element->Name === XlsxWorkbookSerializerCommon::MainNs . 't') {
                $textNodes[] = (string)$element->Value;
            }
        }

        return count($textNodes) > 0 ? implode('', $textNodes) : (string)$inlineStringElement->Value;
    }

    public static function TryParseNumber($rawValue, &$numberValue)
    {
        $rawValueStr = (string)$rawValue;
        if (strpos($rawValueStr, 'e') === false && strpos($rawValueStr, 'E') === false && is_numeric($rawValueStr)) {
            if (strpos($rawValueStr, '.') === false && strpos($rawValueStr, ',') === false) {
                $val = (float)$rawValueStr;
                if ((float)(int)$val === $val && $val >= -2147483648 && $val <= 2147483647) {
                    $numberValue = (int)$val;
                    return true;
                }

                $numberValue = $val;
                return true;
            }
        }

        if (is_numeric($rawValueStr)) {
            $val = (float)$rawValueStr;
            if (abs($val - (int)$val) < PHP_FLOAT_EPSILON && $val >= -2147483648 && $val <= 2147483647) {
                $numberValue = (int)$val;
                return true;
            }

            $numberValue = $val;
            return true;
        }

        $numberValue = null;
        return false;
    }

    public static function AddIssue($diagnostics, $options, $issue)
    {
        $diagnostics->Add($issue);
        if ($options->WarningCallback !== null) {
            $warningInfo = new WarningInfo();
            $warningInfo->Code = $issue->Code;
            $warningInfo->Severity = $issue->Severity;
            $warningInfo->Message = $issue->Message;
            $warningInfo->DataLossRisk = $issue->DataLossRisk;
            $warningInfo->PartUri = $issue->PartUri;
            $warningInfo->SheetName = $issue->SheetName;
            $warningInfo->CellRef = $issue->CellRef;
            $warningInfo->RowIndex = $issue->RowIndex;
            $options->WarningCallback->Warning($warningInfo);
        }
    }

    public static function ResolveWorksheetUri($sheetIndex, $relationshipId, $workbookRelationships, ZipArchive $archive, $diagnostics, $sheetName, $options)
    {
        if (!($relationshipId === null || trim($relationshipId) === '') && array_key_exists($relationshipId, $workbookRelationships)) {
            return $workbookRelationships[$relationshipId];
        }

        $fallbackUri = '/xl/worksheets/sheet' . ($sheetIndex + 1) . '.xml';
        if (self::GetEntry($archive, $fallbackUri) !== null) {
            $issue = new LoadIssue('PKG-R001', DiagnosticSeverity::Recoverable, 'Worksheet relationship metadata was incomplete; the worksheet part was resolved by convention.', true);
            $issue->SheetName = $sheetName;
            $issue->PartUri = $fallbackUri;
            self::AddIssue($diagnostics, $options, $issue);
        }

        return $fallbackUri;
    }

    public static function ResolvePartUri($sourceUri, $target)
    {
        if (str_starts_with($target, '/')) {
            return $target;
        }

        $slashIndex = strrpos($sourceUri, '/');
        $directory = $slashIndex === false ? '/' : substr($sourceUri, 0, $slashIndex + 1);
        $path = $directory . $target;

        $segments = explode('/', $path);
        $resolved = array();
        foreach ($segments as $segment) {
            if ($segment === '' || $segment === '.') {
                continue;
            }

            if ($segment === '..') {
                if (count($resolved) > 0) {
                    array_pop($resolved);
                }
                continue;
            }

            $resolved[] = $segment;
        }

        return '/' . implode('/', $resolved);
    }

    public static function RelationshipPointsTo($candidate, $expectedUri)
    {
        return strcasecmp($candidate, $expectedUri) === 0;
    }

    public static function GetRelationshipTargetKind($candidate)
    {
        if (self::EndsWithIgnoreCase($candidate, '/sharedStrings.xml')) {
            return XlsxWorkbookSerializerCommon::SharedStringsRelationshipType;
        }

        if (self::EndsWithIgnoreCase($candidate, '/styles.xml')) {
            return XlsxWorkbookSerializerCommon::StylesRelationshipType;
        }

        if (stripos($candidate, '/worksheets/') !== false) {
            return XlsxWorkbookSerializerCommon::WorksheetRelationshipType;
        }

        return '';
    }

    public static function LooksLikeDateFormat($formatCode)
    {
        if ($formatCode === null || trim($formatCode) === '') {
            return false;
        }

        if (stripos($formatCode, '[$-F400]') !== false || stripos($formatCode, '[$-F800]') !== false) {
            return true;
        }

        $builder = '';
        $inQuote = false;
        $inBracket = false;
        $length = strlen($formatCode);

        for ($index = 0; $index < $length; $index++) {
            $character = $formatCode[$index];
            if ($character === '"') {
                $inQuote = !$inQuote;
                continue;
            }

            if ($inQuote) {
                continue;
            }

            if ($character === '[') {
                $inBracket = true;
                continue;
            }

            if ($character === ']' && $inBracket) {
                $inBracket = false;
                continue;
            }

            if ($inBracket) {
                $builder .= strtolower($character);
                continue;
            }

            if ($character === '\\' || $character === '_') {
                $index++;
                continue;
            }

            if ($character === '*') {
                continue;
            }

            $builder .= strtolower($character);
        }

        return strpos($builder, 'yy') !== false
            || strpos($builder, 'dd') !== false
            || strpos($builder, 'mm') !== false
            || strpos($builder, 'm/') !== false
            || strpos($builder, '/m') !== false
            || strpos($builder, 'h') !== false
            || strpos($builder, 'am/pm') !== false
            || strpos($builder, 'a/p') !== false
            || strpos($builder, 'ss') !== false;
    }

    private static function EndsWithIgnoreCase($value, $suffix)
    {
        $valueLength = strlen($value);
        $suffixLength = strlen($suffix);
        if ($suffixLength > $valueLength) {
            return false;
        }

        return strcasecmp(substr($value, $valueLength - $suffixLength), $suffix) === 0;
    }
}
