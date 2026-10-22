<?php

namespace Aspose\Cells_FOSS\SupportedFeaturesOpenXmlTests;

use Aspose\Cells_FOSS\BorderStyleType;
use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\HorizontalAlignmentType;
use Aspose\Cells_FOSS\Style;
use Aspose\Cells_FOSS\VerticalAlignmentType;

final class OpenXmlComparisonSupport
{
    const SpreadsheetNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    public static function ReadCellDataWithOpenXmlSdk($workbookPath)
    {
        return self::ReadWithOpenXmlSdk($workbookPath, function ($context) {
            $map = array();
            foreach (self::EnumerateOpenXmlCells($context) as $descriptor) {
                $map[ComparisonValueHelpers::BuildCellKey($descriptor->SheetName, $descriptor->CellReference)] =
                    self::CreateOpenXmlCellSnapshot(
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
                $snapshot = self::CreateOpenXmlStyleSnapshot($descriptor->Cell, $context->Styles)
                    ->CopyWithLocation($descriptor->SheetName, $descriptor->CellReference);
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
            $style->Font->Underline != FontUnderlineType::None,
            (bool) $style->Font->IsStrikeout,
            ComparisonValueHelpers::NormalizeColor($style->Font->Color),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\FillPattern', $style->Pattern),
            ComparisonValueHelpers::NormalizeColor($style->ForegroundColor),
            ComparisonValueHelpers::NormalizeColor($style->BackgroundColor),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\BorderStyleType', $style->Borders->Left->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Left->Color),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\BorderStyleType', $style->Borders->Right->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Right->Color),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\BorderStyleType', $style->Borders->Top->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Top->Color),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\BorderStyleType', $style->Borders->Bottom->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Bottom->Color),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\BorderStyleType', $style->Borders->Diagonal->LineStyle),
            ComparisonValueHelpers::NormalizeColor($style->Borders->Diagonal->Color),
            (bool) $style->Borders->DiagonalUp,
            (bool) $style->Borders->DiagonalDown,
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\HorizontalAlignmentType', $style->HorizontalAlignment),
            ComparisonValueHelpers::EnumName('Aspose\\Cells_FOSS\\VerticalAlignmentType', $style->VerticalAlignment),
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
        $context = self::BuildContext($workbookPath);
        return call_user_func($read, $context);
    }

    private static function BuildContext($workbookPath)
    {
        $context = new OpenXmlReadContext();
        $context->WorkbookPath = $workbookPath;
        $context->Package = new OpenXmlPackage($workbookPath);
        $context->WorkbookDocument = $context->Package->ReadXml('xl/workbook.xml');
        $context->WorkbookRoot = $context->WorkbookDocument->Root;
        $context->WorkbookRelationships = $context->Package->ReadRelationships('xl/workbook.xml');

        if ($context->WorkbookRoot === null) {
            throw new \RuntimeException('Workbook XML is empty.');
        }

        $sheetsElement = $context->WorkbookRoot->Element(self::SpreadsheetNs . 'sheets');
        if ($sheetsElement !== null) {
            foreach ($sheetsElement->Elements(self::SpreadsheetNs . 'sheet') as $sheetElement) {
                $relationshipId = ComparisonValueHelpers::AttributeText($sheetElement, self::RelationshipNs . 'id');
                if ($relationshipId === null || !array_key_exists($relationshipId, $context->WorkbookRelationships)) {
                    continue;
                }

                $target = $context->WorkbookRelationships[$relationshipId]['ResolvedTarget'];
                if ($target === null || $target === '') {
                    continue;
                }

                $name = ComparisonValueHelpers::AttributeText($sheetElement, 'name');
                $context->Sheets[] = array(
                    'Name' => $name === null ? '<unknown>' : $name,
                    'Part' => ltrim($target, '/'),
                );
            }
        }

        $workbookProperties = $context->WorkbookRoot->Element(self::SpreadsheetNs . 'workbookPr');
        $date1904 = ComparisonValueHelpers::ParseBoolAttribute($workbookProperties === null ? null : $workbookProperties->Attribute('date1904'));
        $context->DateSystem = $date1904 ? OpenXmlDateSystem::Mac1904 : OpenXmlDateSystem::Windows1900;

        $context->SharedStrings = self::LoadSharedStrings($context);
        $context->Styles = self::LoadOpenXmlStylesContext($context);

        return $context;
    }

    private static function EnumerateOpenXmlCells($context)
    {
        foreach ($context->Sheets as $sheet) {
            $part = ltrim($sheet['Part'], '/');
            if (!$context->Package->Exists($part)) {
                continue;
            }

            $worksheetDocument = $context->Package->ReadXml($part);
            $worksheetRoot = $worksheetDocument->Root;
            if ($worksheetRoot === null) {
                continue;
            }

            foreach ($worksheetRoot->DescendantsAndSelf() as $node) {
                if ($node->name !== self::SpreadsheetNs . 'c') {
                    continue;
                }

                $normalized = '';
                $reference = ComparisonValueHelpers::AttributeText($node, 'r');
                if (!ComparisonValueHelpers::TryNormalizeCellReference($reference, $normalized)) {
                    continue;
                }

                yield new OpenXmlCellDescriptor($sheet['Name'], $normalized, $node);
            }
        }
    }

    private static function CreateOpenXmlCellSnapshot($sheetName, $cellReference, $cell, $sharedStrings, $dateStyleIndexes, $dateSystem)
    {
        $formula = ComparisonValueHelpers::NormalizeFormulaText(self::CellFormulaText($cell));
        $kind = 'Blank';
        $value = self::ReadOpenXmlValue($cell, $sharedStrings, $dateStyleIndexes, $dateSystem, $kind);
        $effectiveType = $formula === '' ? $kind : 'Formula';
        return new CellSnapshot($sheetName, $cellReference, $effectiveType, ComparisonValueHelpers::NormalizeValue($value), $formula);
    }

    private static function CreateOpenXmlStyleSnapshot($cell, $styles)
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
        $valueElement = $cell->Element(self::SpreadsheetNs . 'v');
        $rawValue = $valueElement === null ? null : $valueElement->Value;
        $styleIndexAttribute = $cell->Attribute('s');
        $styleIndex = $styleIndexAttribute === null ? null : (int) $styleIndexAttribute->Value;
        $isDateStyle = $styleIndex !== null && isset($dateStyleIndexes[$styleIndex]);

        if ($cellType === 'inlineStr') {
            $kind = 'String';
            return self::ReadInlineString($cell->Element(self::SpreadsheetNs . 'is'));
        }

        if ($cellType === 's') {
            $kind = 'String';
            $index = ComparisonValueHelpers::ParseIntValue($rawValue === null ? null : (string) $rawValue);
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
                $kind = 'DateTime';
                return $rawValue;
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
        if (ComparisonValueHelpers::TryParseNumber($rawValue, $numericValue, $kind)) {
            return $numericValue;
        }

        return null;
    }

    private static function CellFormulaText($cell)
    {
        $formulaElement = $cell->Element(self::SpreadsheetNs . 'f');
        return $formulaElement === null ? null : $formulaElement->Value;
    }

    private static function ReadInlineString($inlineString)
    {
        if ($inlineString === null) {
            return '';
        }

        return self::ReadTextElements($inlineString);
    }

    private static function ReadTextElements($element)
    {
        $texts = array();
        foreach ($element->DescendantsAndSelf() as $node) {
            if ($node->name === self::SpreadsheetNs . 't') {
                $texts[] = $node->Value;
            }
        }

        if (count($texts) === 0) {
            return $element->Value;
        }

        return implode('', $texts);
    }

    private static function LoadSharedStrings($context)
    {
        $part = self::ResolveRelationshipPart($context, '/sharedStrings');
        if ($part === null) {
            return array();
        }

        $document = $context->Package->ReadXml($part);
        $root = $document->Root;
        if ($root === null) {
            return array();
        }

        $strings = array();
        foreach ($root->Elements(self::SpreadsheetNs . 'si') as $item) {
            $strings[] = self::ReadTextElements($item);
        }

        return $strings;
    }

    private static function ResolveRelationshipPart($context, $typeFragment)
    {
        foreach ($context->WorkbookRelationships as $relationship) {
            if (strpos($relationship['Type'], $typeFragment) === false) {
                continue;
            }

            $target = $relationship['ResolvedTarget'];
            if ($target === null || $target === '') {
                continue;
            }

            return ltrim($target, '/');
        }

        return null;
    }

    private static function LoadOpenXmlStylesContext($context)
    {
        $part = self::ResolveRelationshipPart($context, '/styles');
        if ($part === null || !$context->Package->Exists($part)) {
            return OpenXmlStylesContext::Default();
        }

        $document = $context->Package->ReadXml($part);
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
        $cellXfsElement = $root->Element(self::SpreadsheetNs . 'cellXfs');
        $cellFormats = $cellXfsElement === null ? array() : $cellXfsElement->Elements(self::SpreadsheetNs . 'xf');
        if (count($cellFormats) === 0) {
            $styleSnapshots[] = StyleSnapshot::Default();
            return new OpenXmlStylesContext($styleSnapshots, $dateStyleIndexes);
        }

        for ($index = 0; $index < count($cellFormats); $index++) {
            $cellFormat = $cellFormats[$index];
            $fontId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('fontId'));
            $fillId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('fillId'));
            $borderId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('borderId'));
            $numberFormatId = ComparisonValueHelpers::ParseIntAttribute($cellFormat->Attribute('numFmtId'));
            if ($fontId === null) {
                $fontId = 0;
            }
            if ($fillId === null) {
                $fillId = 0;
            }
            if ($borderId === null) {
                $borderId = 0;
            }
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
                self::AlignmentInt($alignmentElement, 'indent'),
                self::AlignmentInt($alignmentElement, 'textRotation'),
                ComparisonValueHelpers::ParseBoolAttribute($alignmentElement === null ? null : $alignmentElement->Attribute('shrinkToFit')),
                self::AlignmentInt($alignmentElement, 'readingOrder'),
                self::AlignmentInt($alignmentElement, 'relativeIndent'),
                $numberFormatId,
                $customFormatCode,
                self::ProtectionLocked($protectionElement),
                ComparisonValueHelpers::ParseBoolAttribute($protectionElement === null ? null : $protectionElement->Attribute('hidden'))
            );

            if (ComparisonValueHelpers::IsBuiltInDateFormat($numberFormatId)
                || ($customFormatCode !== '' && ComparisonValueHelpers::LooksLikeDateFormat($customFormatCode))) {
                $dateStyleIndexes[$index] = true;
            }
        }

        return new OpenXmlStylesContext($styleSnapshots, $dateStyleIndexes);
    }

    private static function AlignmentInt($alignmentElement, $name)
    {
        if ($alignmentElement === null) {
            return 0;
        }

        $value = ComparisonValueHelpers::ParseIntAttribute($alignmentElement->Attribute($name));
        return $value === null ? 0 : $value;
    }

    private static function ProtectionLocked($protectionElement)
    {
        if ($protectionElement === null) {
            return true;
        }

        $value = ComparisonValueHelpers::ParseOptionalBoolAttribute($protectionElement->Attribute('locked'));
        return $value === null ? true : $value;
    }

    private static function ReadOpenXmlFont($font)
    {
        $nameElement = $font->Element(self::SpreadsheetNs . 'name');
        $sizeElement = $font->Element(self::SpreadsheetNs . 'sz');

        return new OpenXmlFontDescriptor(
            $nameElement === null || ComparisonValueHelpers::AttributeText($nameElement, 'val') === null
                ? 'Calibri'
                : ComparisonValueHelpers::AttributeText($nameElement, 'val'),
            ComparisonValueHelpers::NormalizeDouble(ComparisonValueHelpers::ParseDoubleValue(
                $sizeElement === null ? null : ComparisonValueHelpers::AttributeText($sizeElement, 'val'),
                11.0
            )),
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
