<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\BorderStyleType;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\HorizontalAlignmentType;
use Aspose\Cells_FOSS\Style;
use Aspose\Cells_FOSS\VerticalAlignmentType;

final class OpenXmlComparisonSupport
{
    const SpreadsheetNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipsNs = '{http://schemas.openxmlformats.org/package/2006/relationships}';
    const RelationshipNsId = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id';

    public static function ReadCellDataWithOpenXmlSdk($workbookPath)
    {
        return self::ReadWithOpenXmlSdk($workbookPath, function ($context) {
            $map = array();
            foreach (self::EnumerateOpenXmlCells($context) as $descriptor) {
                $map[ComparisonValueHelpers::BuildCellKey($descriptor->SheetName, $descriptor->CellReference)] = self::CreateOpenXmlCellSnapshot(
                    $descriptor->SheetName,
                    $descriptor->CellReference,
                    $descriptor->Cell,
                    $context->SharedStrings,
                    $context->Styles->DateStyleIndexes,
                    $context->DateSystem
                );
            }

            return $map;
        });
    }

    public static function ReadCellStylesWithOpenXmlSdk($workbookPath)
    {
        return self::ReadWithOpenXmlSdk($workbookPath, function ($context) {
            $map = array();
            foreach (self::EnumerateOpenXmlCells($context) as $descriptor) {
                $snapshot = self::CreateOpenXmlStyleSnapshot($descriptor->Cell, $context->Styles)->CopyWithLocation($descriptor->SheetName, $descriptor->CellReference);
                $map[ComparisonValueHelpers::BuildCellKey($descriptor->SheetName, $descriptor->CellReference)] = $snapshot;
            }

            return $map;
        });
    }

    public static function CreateLibraryStyleSnapshot($sheetName, $cellReference, Style $style)
    {
        return new StyleSnapshot(
            $sheetName,
            $cellReference,
            $style->Font->Name,
            ComparisonValueHelpers::NormalizeDouble($style->Font->Size),
            (bool) $style->Font->IsBold,
            (bool) $style->Font->IsItalic,
            $style->Font->Underline !== FontUnderlineType::None,
            (bool) $style->Font->IsStrikeout,
            ComparisonValueHelpers::NormalizeColor($style->Font->Color),
            ComparisonValueHelpers::EnumName(FillPattern::class, $style->Pattern),
            ComparisonValueHelpers::NormalizeColor($style->ForegroundColor),
            ComparisonValueHelpers::NormalizeColor($style->BackgroundColor),
            ComparisonValueHelpers::EnumName(BorderStyleType::class, $style->Borders->Left->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Left->Color),
            ComparisonValueHelpers::EnumName(BorderStyleType::class, $style->Borders->Right->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Right->Color),
            ComparisonValueHelpers::EnumName(BorderStyleType::class, $style->Borders->Top->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Top->Color),
            ComparisonValueHelpers::EnumName(BorderStyleType::class, $style->Borders->Bottom->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Bottom->Color),
            ComparisonValueHelpers::EnumName(BorderStyleType::class, $style->Borders->Diagonal->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Diagonal->Color),
            (bool) $style->Borders->DiagonalUp,
            (bool) $style->Borders->DiagonalDown,
            ComparisonValueHelpers::EnumName(HorizontalAlignmentType::class, $style->HorizontalAlignment),
            ComparisonValueHelpers::EnumName(VerticalAlignmentType::class, $style->VerticalAlignment),
            (bool) $style->WrapText,
            $style->IndentLevel === null ? 0 : $style->IndentLevel,
            $style->TextRotation === null ? 0 : $style->TextRotation,
            (bool) $style->ShrinkToFit,
            $style->ReadingOrder === null ? 0 : $style->ReadingOrder,
            $style->RelativeIndent === null ? 0 : $style->RelativeIndent,
            $style->Number === null ? 0 : $style->Number,
            $style->Custom === null ? '' : $style->Custom,
            $style->IsLocked === null ? true : (bool) $style->IsLocked,
            (bool) $style->IsHidden
        );
    }

    private static function ReadWithOpenXmlSdk($workbookPath, $read)
    {
        $archive = new \ZipArchive();
        if ($archive->open($workbookPath) !== true) {
            throw new \RuntimeException('WorkbookPart is missing.');
        }

        try {
            $context = self::CreateOpenXmlReadContext($archive);
            return call_user_func($read, $context);
        } finally {
            $archive->close();
        }
    }

    private static function CreateOpenXmlReadContext($archive)
    {
        if (!self::EntryExists($archive, 'xl/workbook.xml')) {
            throw new \RuntimeException('WorkbookPart is missing.');
        }

        $workbookDocument = XDocument::Load(self::ReadEntry($archive, 'xl/workbook.xml'));
        $workbookRoot = $workbookDocument->Root;
        if ($workbookRoot === null) {
            throw new \RuntimeException('WorkbookPart is missing.');
        }

        $sharedStrings = self::LoadSharedStrings($archive);
        $workbookProperties = $workbookRoot->Element(self::SpreadsheetNs . 'workbookPr');
        $date1904Attribute = $workbookProperties === null ? null : $workbookProperties->Attribute('date1904');
        $dateSystem = ComparisonValueHelpers::ParseBoolAttribute($date1904Attribute) ? OpenXmlDateSystem::Mac1904 : OpenXmlDateSystem::Windows1900;
        $styles = self::LoadOpenXmlStylesContext($archive);
        $sheets = self::ResolveSheets($archive, $workbookRoot);

        return new OpenXmlReadContext($archive, $sharedStrings, $dateSystem, $styles, $sheets);
    }

    private static function ResolveSheets($archive, $workbookRoot)
    {
        $sheetsElement = $workbookRoot->Element(self::SpreadsheetNs . 'sheets');
        $sheetElements = $sheetsElement === null ? array() : $sheetsElement->Elements(self::SpreadsheetNs . 'sheet');
        $relationships = self::ReadWorkbookRelationships($archive);

        $sheets = array();
        foreach ($sheetElements as $sheet) {
            $relationshipId = ComparisonValueHelpers::AttributeText($sheet, self::RelationshipNsId);
            if ($relationshipId === null) {
                continue;
            }

            if (!array_key_exists($relationshipId, $relationships)) {
                continue;
            }

            $partPath = self::ResolvePart('xl/workbook.xml', $relationships[$relationshipId]);
            if ($partPath === null) {
                continue;
            }

            $name = ComparisonValueHelpers::AttributeText($sheet, 'name');
            if ($name === null) {
                $name = '<unknown>';
            }

            $sheets[] = array('Name' => $name, 'Path' => $partPath);
        }

        return $sheets;
    }

    private static function ReadWorkbookRelationships($archive)
    {
        $result = array();
        if (!self::EntryExists($archive, 'xl/_rels/workbook.xml.rels')) {
            return $result;
        }

        $document = XDocument::Load(self::ReadEntry($archive, 'xl/_rels/workbook.xml.rels'));
        $root = $document->Root;
        if ($root === null) {
            return $result;
        }

        foreach ($root->Elements(self::RelationshipsNs . 'Relationship') as $relationship) {
            $id = ComparisonValueHelpers::AttributeText($relationship, 'Id');
            if ($id === null) {
                continue;
            }

            $target = ComparisonValueHelpers::AttributeText($relationship, 'Target');
            $result[$id] = $target === null ? '' : $target;
        }

        return $result;
    }

    private static function ResolvePart($sourcePart, $relationshipTarget)
    {
        if ($relationshipTarget === null) {
            return null;
        }

        $target = str_replace('\\', '/', trim($relationshipTarget));
        if ($target === '') {
            return null;
        }

        if (strpos($target, '/') === 0) {
            return trim(self::NormalizePart(ltrim($target, '/')), '/');
        }

        $baseDirectory = '';
        $slashIndex = strrpos($sourcePart, '/');
        if ($slashIndex !== false) {
            $baseDirectory = substr($sourcePart, 0, $slashIndex + 1);
        }

        return trim(self::NormalizePart($baseDirectory . $target), '/');
    }

    private static function NormalizePart($path)
    {
        $segments = explode('/', str_replace('\\', '/', $path));
        $stack = array();
        foreach ($segments as $segment) {
            if ($segment === '' || $segment === '.') {
                continue;
            }

            if ($segment === '..') {
                if (count($stack) > 0) {
                    array_pop($stack);
                }

                continue;
            }

            $stack[] = $segment;
        }

        return implode('/', $stack);
    }

    private static function ReadEntry($archive, $path)
    {
        $normalized = ltrim(str_replace('\\', '/', $path), '/');
        $index = $archive->locateName($normalized);
        if ($index === false) {
            return null;
        }

        return $archive->getFromIndex($index);
    }

    private static function EntryExists($archive, $path)
    {
        $normalized = ltrim(str_replace('\\', '/', $path), '/');
        return $archive->locateName($normalized) !== false;
    }

    private static function EnumerateOpenXmlCells(OpenXmlReadContext $context)
    {
        $descriptors = array();
        foreach ($context->Sheets as $sheet) {
            $worksheetXml = self::ReadEntry($context->Archive, $sheet['Path']);
            if ($worksheetXml === null) {
                continue;
            }

            $document = XDocument::Load($worksheetXml);
            $root = $document->Root;
            if ($root === null) {
                continue;
            }

            foreach (self::Descendants($root) as $element) {
                if ($element->name !== self::SpreadsheetNs . 'c') {
                    continue;
                }

                $cellReference = '';
                if (!ComparisonValueHelpers::TryNormalizeCellReference(ComparisonValueHelpers::AttributeText($element, 'r'), $cellReference)) {
                    continue;
                }

                $descriptors[] = new OpenXmlCellDescriptor($sheet['Name'], $cellReference, $element);
            }
        }

        return $descriptors;
    }

    private static function Descendants($element)
    {
        $result = array();
        foreach ($element->nodes as $node) {
            if ($node instanceof XElement) {
                $result[] = $node;
                foreach (self::Descendants($node) as $descendant) {
                    $result[] = $descendant;
                }
            }
        }

        return $result;
    }

    private static function CreateOpenXmlCellSnapshot($sheetName, $cellReference, $cell, $sharedStrings, $dateStyleIndexes, $dateSystem)
    {
        $formula = ComparisonValueHelpers::NormalizeFormulaText(self::ReadChildText($cell, 'f'));
        $kind = 'Blank';
        $value = self::ReadOpenXmlValue($cell, $sharedStrings, $dateStyleIndexes, $dateSystem, $kind);
        $effectiveType = $formula === '' ? $kind : 'Formula';
        return new CellSnapshot($sheetName, $cellReference, $effectiveType, ComparisonValueHelpers::NormalizeValue($value), $formula);
    }

    private static function CreateOpenXmlStyleSnapshot($cell, OpenXmlStylesContext $styles)
    {
        $styleIndexAttribute = $cell->Attribute('s');
        $styleIndex = $styleIndexAttribute === null ? 0 : (int) $styleIndexAttribute->Value;
        if ($styleIndex < 0 || $styleIndex >= count($styles->CellFormats)) {
            $styleIndex = 0;
        }

        return $styles->CellFormats[$styleIndex];
    }

    private static function ReadOpenXmlValue($cell, $sharedStrings, $dateStyleIndexes, $dateSystem, &$kind)
    {
        $kind = 'Blank';
        $cellType = ComparisonValueHelpers::AttributeText($cell, 't');
        $rawValue = self::ReadChildText($cell, 'v');
        $styleIndexAttribute = $cell->Attribute('s');
        $styleIndex = $styleIndexAttribute === null ? null : (int) $styleIndexAttribute->Value;
        $isDateStyle = $styleIndex !== null && in_array($styleIndex, $dateStyleIndexes, true);

        if ($cellType === 'inlineStr') {
            $kind = 'String';
            return self::ReadInlineString($cell->Element(self::SpreadsheetNs . 'is'));
        }

        if ($cellType === 's') {
            $kind = 'String';
            $index = ComparisonValueHelpers::ParseIntValue($rawValue);
            if ($index !== null && $index >= 0 && $index < count($sharedStrings)) {
                return $sharedStrings[$index];
            }

            return '';
        }

        if ($cellType === 'b') {
            $kind = 'Boolean';
            return $rawValue === '1' || ($rawValue !== null && strcasecmp($rawValue, 'true') === 0);
        }

        if ($cellType === 'd') {
            if ($rawValue !== null && trim($rawValue) !== '') {
                try {
                    $dateValue = new \DateTime($rawValue);
                    $kind = 'DateTime';
                    return $dateValue;
                } catch (\Exception $exception) {
                    return null;
                }
            }

            return null;
        }

        if ($cellType === 'str' || $cellType === 'e') {
            $kind = 'String';
            return $rawValue === null ? '' : $rawValue;
        }

        if ($rawValue === null || $rawValue === '') {
            return null;
        }

        if ($isDateStyle) {
            if (is_numeric($rawValue)) {
                $kind = 'DateTime';
                return self::ConvertSerialToDateTime((float) $rawValue, $dateSystem);
            }

            return null;
        }

        $numericValue = null;
        $numericKind = 'Blank';
        if (ComparisonValueHelpers::TryParseNumber($rawValue, $numericValue, $numericKind)) {
            $kind = $numericKind;
            return $numericValue;
        }

        return null;
    }

    private static function ReadChildText($element, $localName)
    {
        $child = $element->Element(self::SpreadsheetNs . $localName);
        if ($child === null) {
            return null;
        }

        return $child->Value;
    }

    private static function ReadInlineString($inlineString)
    {
        if ($inlineString === null) {
            return '';
        }

        $texts = array();
        foreach (self::Descendants($inlineString) as $node) {
            if ($node->name === self::SpreadsheetNs . 't') {
                $texts[] = $node->Value;
            }
        }

        return count($texts) === 0 ? $inlineString->Value : implode('', $texts);
    }

    private static function LoadSharedStrings($archive)
    {
        $xml = self::ReadEntry($archive, 'xl/sharedStrings.xml');
        if ($xml === null) {
            return array();
        }

        $document = XDocument::Load($xml);
        $root = $document->Root;
        if ($root === null) {
            return array();
        }

        $result = array();
        foreach ($root->Elements(self::SpreadsheetNs . 'si') as $item) {
            $texts = array();
            foreach (self::Descendants($item) as $node) {
                if ($node->name === self::SpreadsheetNs . 't') {
                    $texts[] = $node->Value;
                }
            }

            $result[] = count($texts) === 0 ? $item->Value : implode('', $texts);
        }

        return $result;
    }

    private static function LoadOpenXmlStylesContext($archive)
    {
        if (!self::EntryExists($archive, 'xl/styles.xml')) {
            return OpenXmlStylesContext::Default();
        }

        $document = XDocument::Load(self::ReadEntry($archive, 'xl/styles.xml'));
        $root = $document->Root;
        if ($root === null) {
            return OpenXmlStylesContext::Default();
        }

        $customFormats = array();
        $numFmtsElement = $root->Element(self::SpreadsheetNs . 'numFmts');
        if ($numFmtsElement !== null) {
            foreach ($numFmtsElement->Elements(self::SpreadsheetNs . 'numFmt') as $numFmt) {
                $id = ComparisonValueHelpers::ParseIntAttribute($numFmt->Attribute('numFmtId'));
                $formatCode = ComparisonValueHelpers::AttributeText($numFmt, 'formatCode');
                if ($id !== null && $formatCode !== null && $formatCode !== '') {
                    $customFormats[$id] = $formatCode;
                }
            }
        }

        $fonts = array();
        $fontsElement = $root->Element(self::SpreadsheetNs . 'fonts');
        if ($fontsElement !== null) {
            foreach ($fontsElement->Elements(self::SpreadsheetNs . 'font') as $font) {
                $fonts[] = self::ReadOpenXmlFont($font);
            }
        }

        if (count($fonts) === 0) {
            $fonts[] = OpenXmlFontDescriptor::Default();
        }

        $fills = array();
        $fillsElement = $root->Element(self::SpreadsheetNs . 'fills');
        if ($fillsElement !== null) {
            foreach ($fillsElement->Elements(self::SpreadsheetNs . 'fill') as $fill) {
                $fills[] = self::ReadOpenXmlFill($fill);
            }
        }

        if (count($fills) === 0) {
            $fills[] = OpenXmlFillDescriptor::Default();
        }

        $borders = array();
        $bordersElement = $root->Element(self::SpreadsheetNs . 'borders');
        if ($bordersElement !== null) {
            foreach ($bordersElement->Elements(self::SpreadsheetNs . 'border') as $border) {
                $borders[] = self::ReadOpenXmlBorder($border);
            }
        }

        if (count($borders) === 0) {
            $borders[] = OpenXmlBorderDescriptor::Default();
        }

        $styleSnapshots = array();
        $dateStyleIndexes = array();
        $cellFormatsElement = $root->Element(self::SpreadsheetNs . 'cellXfs');
        $cellFormats = $cellFormatsElement === null ? array() : $cellFormatsElement->Elements(self::SpreadsheetNs . 'xf');
        if (count($cellFormats) === 0) {
            $styleSnapshots[] = StyleSnapshot::Default();
            return new OpenXmlStylesContext($styleSnapshots, $dateStyleIndexes);
        }

        for ($index = 0; $index < count($cellFormats); $index++) {
            $cellFormat = $cellFormats[$index];
            $fontId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('fontId'));
            if ($fontId === null) {
                $fontId = 0;
            }

            $fillId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('fillId'));
            if ($fillId === null) {
                $fillId = 0;
            }

            $borderId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('borderId'));
            if ($borderId === null) {
                $borderId = 0;
            }

            $numberFormatId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('numFmtId'));
            if ($numberFormatId === null) {
                $numberFormatId = 0;
            }

            $font = ComparisonValueHelpers::GetOrDefault($fonts, $fontId, OpenXmlFontDescriptor::Default());
            $fill = ComparisonValueHelpers::GetOrDefault($fills, $fillId, OpenXmlFillDescriptor::Default());
            $border = ComparisonValueHelpers::GetOrDefault($borders, $borderId, OpenXmlBorderDescriptor::Default());
            $alignmentElement = $cellFormat->Element(self::SpreadsheetNs . 'alignment');
            $protectionElement = $cellFormat->Element(self::SpreadsheetNs . 'protection');
            $customFormatCode = array_key_exists($numberFormatId, $customFormats) ? $customFormats[$numberFormatId] : '';

            $styleSnapshots[] = new StyleSnapshot(
                '',
                '',
                $font->Name,
                $font->Size,
                $font->Bold,
                $font->Italic,
                $font->Underline,
                $font->StrikeThrough,
                $font->Color,
                $fill->Pattern,
                $fill->ForegroundColor,
                $fill->BackgroundColor,
                $border->Left->Style,
                $border->Left->Color,
                $border->Right->Style,
                $border->Right->Color,
                $border->Top->Style,
                $border->Top->Color,
                $border->Bottom->Style,
                $border->Bottom->Color,
                $border->Diagonal->Style,
                $border->Diagonal->Color,
                $border->DiagonalUp,
                $border->DiagonalDown,
                ComparisonValueHelpers::ParseHorizontalAlignment($alignmentElement === null ? null : ComparisonValueHelpers::AttributeText($alignmentElement, 'horizontal')),
                ComparisonValueHelpers::ParseVerticalAlignment($alignmentElement === null ? null : ComparisonValueHelpers::AttributeText($alignmentElement, 'vertical')),
                ComparisonValueHelpers::ParseBoolAttribute($alignmentElement === null ? null : $alignmentElement->Attribute('wrapText')),
                self::OptionalInt($alignmentElement, 'indent'),
                self::OptionalInt($alignmentElement, 'textRotation'),
                ComparisonValueHelpers::ParseBoolAttribute($alignmentElement === null ? null : $alignmentElement->Attribute('shrinkToFit')),
                self::OptionalInt($alignmentElement, 'readingOrder'),
                self::OptionalInt($alignmentElement, 'relativeIndent'),
                $numberFormatId,
                $customFormatCode,
                self::OptionalBool($protectionElement, 'locked', true),
                ComparisonValueHelpers::ParseBoolAttribute($protectionElement === null ? null : $protectionElement->Attribute('hidden'))
            );

            if (ComparisonValueHelpers::IsBuiltInDateFormat($numberFormatId) || ($customFormatCode !== '' && ComparisonValueHelpers::LooksLikeDateFormat($customFormatCode))) {
                $dateStyleIndexes[] = $index;
            }
        }

        return new OpenXmlStylesContext($styleSnapshots, $dateStyleIndexes);
    }

    private static function OptionalInt($element, $attributeName)
    {
        if ($element === null) {
            return 0;
        }

        $value = ComparisonValueHelpers::ParseIntAttribute($element->Attribute($attributeName));
        return $value === null ? 0 : $value;
    }

    private static function OptionalBool($element, $attributeName, $default)
    {
        if ($element === null) {
            return $default;
        }

        $value = ComparisonValueHelpers::ParseOptionalBoolAttribute($element->Attribute($attributeName));
        return $value === null ? $default : $value;
    }

    private static function ReadOpenXmlFont($font)
    {
        $nameElement = $font->Element(self::SpreadsheetNs . 'name');
        $name = $nameElement === null ? null : ComparisonValueHelpers::AttributeText($nameElement, 'val');
        if ($name === null) {
            $name = 'Calibri';
        }

        $sizeElement = $font->Element(self::SpreadsheetNs . 'sz');
        $sizeText = $sizeElement === null ? null : ComparisonValueHelpers::AttributeText($sizeElement, 'val');

        return new OpenXmlFontDescriptor(
            $name,
            ComparisonValueHelpers::NormalizeDouble(ComparisonValueHelpers::ParseDoubleValue($sizeText, 11.0)),
            $font->Element(self::SpreadsheetNs . 'b') !== null,
            $font->Element(self::SpreadsheetNs . 'i') !== null,
            $font->Element(self::SpreadsheetNs . 'u') !== null,
            $font->Element(self::SpreadsheetNs . 'strike') !== null,
            ComparisonValueHelpers::NormalizeOpenXmlColor($font->Element(self::SpreadsheetNs . 'color'))
        );
    }

    private static function ReadOpenXmlFill($fill)
    {
        $patternFill = $fill->Element(self::SpreadsheetNs . 'patternFill');
        if ($patternFill === null) {
            return OpenXmlFillDescriptor::Default();
        }

        return new OpenXmlFillDescriptor(
            ComparisonValueHelpers::ParseFillPattern(ComparisonValueHelpers::AttributeText($patternFill, 'patternType')),
            ComparisonValueHelpers::NormalizeOpenXmlColor($patternFill->Element(self::SpreadsheetNs . 'fgColor')),
            ComparisonValueHelpers::NormalizeOpenXmlColor($patternFill->Element(self::SpreadsheetNs . 'bgColor'))
        );
    }

    private static function ReadOpenXmlBorder($border)
    {
        return new OpenXmlBorderDescriptor(
            self::ReadOpenXmlBorderSide($border->Element(self::SpreadsheetNs . 'left')),
            self::ReadOpenXmlBorderSide($border->Element(self::SpreadsheetNs . 'right')),
            self::ReadOpenXmlBorderSide($border->Element(self::SpreadsheetNs . 'top')),
            self::ReadOpenXmlBorderSide($border->Element(self::SpreadsheetNs . 'bottom')),
            self::ReadOpenXmlBorderSide($border->Element(self::SpreadsheetNs . 'diagonal')),
            ComparisonValueHelpers::ParseBoolAttribute($border->Attribute('diagonalUp')),
            ComparisonValueHelpers::ParseBoolAttribute($border->Attribute('diagonalDown'))
        );
    }

    private static function ReadOpenXmlBorderSide($side)
    {
        if ($side === null) {
            return OpenXmlBorderSideDescriptor::Default();
        }

        return new OpenXmlBorderSideDescriptor(
            ComparisonValueHelpers::ParseBorderStyle(ComparisonValueHelpers::AttributeText($side, 'style')),
            ComparisonValueHelpers::NormalizeOpenXmlColor($side->Element(self::SpreadsheetNs . 'color'))
        );
    }

    private static function ConvertSerialToDateTime($serial, $dateSystem)
    {
        $libraryDateSystem = $dateSystem === OpenXmlDateSystem::Mac1904 ? DateSystem::Mac1904 : DateSystem::Windows1900;
        return DateSerialConverter::FromSerial($serial, $libraryDateSystem);
    }
}
