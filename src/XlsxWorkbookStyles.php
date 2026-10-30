<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AlignmentValue;
use Aspose\Cells_FOSS\Core\BorderSideValue;
use Aspose\Cells_FOSS\Core\BordersValue;
use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellRecord;
use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Core\NumberFormatValue;
use Aspose\Cells_FOSS\Core\ProtectionValue;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;

final class XlsxWorkbookStyles
{
    private function __construct()
    {
    }

    public static function BuildStylesheet(WorkbookModel $model)
    {
        $defaultStyle = $model->DefaultStyle->Clone();

        $fonts = array($defaultStyle->Font->Clone());
        $fontIds = array();
        $fontIds[self::GetFontKey($fonts[0])] = 0;

        $fill0 = self::ToFillValue($defaultStyle);
        $fill1 = new FillValue();
        $fill1->Pattern = FillPatternKind::Gray125;
        $fills = array($fill0, $fill1);
        $fillIds = array();
        $fillIds[self::GetFillKey($fills[0])] = 0;
        $fillIds[self::GetFillKey($fills[1])] = 1;

        $borders = array($defaultStyle->Borders->Clone());
        $borderIds = array();
        $borderIds[self::GetBordersKey($borders[0])] = 0;

        $customNumberFormatIds = array();
        $customNumberFormats = array();
        $nextCustomNumberFormatId = 164;
        $defaultNumberFormatId = self::ResolveNumberFormatId($defaultStyle->NumberFormat, $customNumberFormatIds, $customNumberFormats, $nextCustomNumberFormatId);

        $normalCellFormat = new CellFormatValue();
        $normalCellFormat->NumFmtId = $defaultNumberFormatId;
        $normalCellFormat->FontId = 0;
        $normalCellFormat->FillId = 0;
        $normalCellFormat->BorderId = 0;
        $normalCellFormat->Alignment = $defaultStyle->Alignment->Clone();
        $normalCellFormat->Protection = $defaultStyle->Protection->Clone();

        $cellFormats = array($normalCellFormat);
        $styleIndices = array();
        $styleIndices[self::GetStyleKey($defaultStyle)] = 0;
        $differentialFormats = array();
        $differentialStyleIndices = array();

        foreach ($model->Worksheets as $worksheet) {
            $persistedEntries = self::CollectPersistedCellEntries($worksheet, $model->DefaultStyle);
            for ($index = 0; $index < count($persistedEntries); $index++) {
                $pair = $persistedEntries[$index];
                $style = self::GetStyleForSerialization($pair[1]);
                $styleKey = self::GetStyleKey($style);
                if (array_key_exists($styleKey, $styleIndices)) {
                    continue;
                }

                $fontId = self::Intern($fontIds, $fonts, $style->Font, array(__CLASS__, 'GetFontKey'));
                $fillId = self::Intern($fillIds, $fills, self::ToFillValue($style), array(__CLASS__, 'GetFillKey'));
                $borderId = self::Intern($borderIds, $borders, $style->Borders, array(__CLASS__, 'GetBordersKey'));
                $numFmtId = self::ResolveNumberFormatId($style->NumberFormat, $customNumberFormatIds, $customNumberFormats, $nextCustomNumberFormatId);

                $styleIndices[$styleKey] = count($cellFormats);
                $cellFormat = new CellFormatValue();
                $cellFormat->NumFmtId = $numFmtId;
                $cellFormat->FontId = $fontId;
                $cellFormat->FillId = $fillId;
                $cellFormat->BorderId = $borderId;
                $cellFormat->Alignment = $style->Alignment->Clone();
                $cellFormat->Protection = $style->Protection->Clone();
                $cellFormats[] = $cellFormat;
            }

            for ($formattingIndex = 0; $formattingIndex < count($worksheet->ConditionalFormattings); $formattingIndex++) {
                $formatting = $worksheet->ConditionalFormattings[$formattingIndex];
                for ($conditionIndex = 0; $conditionIndex < count($formatting->Conditions); $conditionIndex++) {
                    $style = $formatting->Conditions[$conditionIndex]->Style;
                    if (self::StylesEqual($style, StyleValue::Default())) {
                        continue;
                    }

                    $styleKey = self::GetStyleKey($style);
                    if (array_key_exists($styleKey, $differentialStyleIndices)) {
                        continue;
                    }

                    $differentialStyleIndices[$styleKey] = count($differentialFormats);
                    $differentialFormats[] = $style->Clone();
                }
            }
        }

        $hasStyles = count($styleIndices) > 1 || !self::StylesEqual($defaultStyle, StyleValue::Default()) || count($differentialFormats) > 0;
        $document = XlsxWorkbookStylesXml::BuildStylesheetDocument($fonts, $fills, $borders, $normalCellFormat, $cellFormats, $customNumberFormats, $differentialFormats);
        return new StylesheetSaveContext($document, $styleIndices, $differentialStyleIndices, count($differentialFormats), $hasStyles);
    }

    public static function GetStyleForSerialization(CellRecord $record)
    {
        $style = self::GetEffectiveStyle($record);
        if ($record->Kind !== CellValueKind::DateTime || self::IsDateNumberFormat($style->NumberFormat)) {
            return $style;
        }

        $serializedStyle = $style->Clone();
        $numberFormat = new NumberFormatValue();
        $numberFormat->Number = 14;
        $numberFormat->Custom = null;
        $serializedStyle->NumberFormat = $numberFormat;
        return $serializedStyle;
    }

    private static function CollectPersistedCellEntries(WorksheetModel $worksheet, StyleValue $workbookDefaultStyle)
    {
        $persistedEntries = array();
        foreach ($worksheet->Cells as $key => $record) {
            if (XlsxWorkbookSerializerCommon::ShouldPersistCell($workbookDefaultStyle, $record)) {
                $persistedEntries[] = array(CellAddress::Parse($key), $record);
            }
        }

        usort($persistedEntries, array(__CLASS__, 'CompareCellEntries'));
        return $persistedEntries;
    }

    private static function CompareCellEntries($left, $right)
    {
        $rowComparison = $left[0]->RowIndex <=> $right[0]->RowIndex;
        if ($rowComparison !== 0) {
            return $rowComparison;
        }

        return $left[0]->ColumnIndex <=> $right[0]->ColumnIndex;
    }

    public static function LoadStylesheet(ZipArchive $archive, $workbookRelationships, $options, $diagnostics)
    {
        $stylesUri = XlsxWorkbookArchiveHelpers::FindRelationshipTarget($workbookRelationships, '/xl/styles.xml');
        if ($stylesUri === '' || $stylesUri === null) {
            if (count($workbookRelationships) === 0) {
                $stylesUri = '/xl/styles.xml';
            } else {
                $stylesUri = XlsxWorkbookArchiveHelpers::FindRelationshipTargetByType($workbookRelationships, XlsxWorkbookSerializerCommon::StylesRelationshipType);
            }
        }

        if ($stylesUri === '' || $stylesUri === null) {
            $stylesUri = '/xl/styles.xml';
        }

        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $stylesUri);
        if ($entry === null) {
            return new StylesheetLoadContext();
        }

        $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        $root = $document->Root;
        if ($root === null) {
            return new StylesheetLoadContext();
        }

        $customFormats = array();
        $numFmtsElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'numFmts');
        if ($numFmtsElement !== null) {
            foreach ($numFmtsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'numFmt') as $numFmt) {
                $id = XlsxWorkbookArchiveHelpers::ParseIntAttribute($numFmt->Attribute('numFmtId'));
                $code = (string)$numFmt->Attribute('formatCode');
                if ($id !== null && !($code === null || $code === '')) {
                    $customFormats[$id] = $code;
                }
            }
        }

        $fonts = XlsxWorkbookStylesXml::ReadFontValues($root);
        if (count($fonts) === 0) {
            $fonts[] = StyleValue::Default()->Font->Clone();
        }

        $fills = XlsxWorkbookStylesXml::ReadFillValues($root);
        if (count($fills) === 0) {
            $fillNone = new FillValue();
            $fillNone->Pattern = FillPatternKind::None;
            $fills[] = $fillNone;
            $fillGray = new FillValue();
            $fillGray->Pattern = FillPatternKind::Gray125;
            $fills[] = $fillGray;
        }

        $borders = XlsxWorkbookStylesXml::ReadBordersValues($root);
        if (count($borders) === 0) {
            $borders[] = StyleValue::Default()->Borders->Clone();
        }

        $context = new StylesheetLoadContext();
        $context->CellFormats = array();
        $differentialFormats = XlsxWorkbookStylesXml::ReadDifferentialStyleValues($root);
        foreach ($differentialFormats as $differentialFormat) {
            $context->DifferentialFormats[] = $differentialFormat;
        }

        $tableStyles = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'tableStyles');
        $context->RawTableStylesXml = $tableStyles === null ? null : $tableStyles->toString();
        $differentialFormatsElement = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'dxfs');
        if ($differentialFormatsElement !== null) {
            foreach ($differentialFormatsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'dxf') as $differentialFormatElement) {
                $context->RawDifferentialFormatsXml[] = $differentialFormatElement->toString();
            }
        }

        $cellXfs = array();
        $cellXfsContainer = $root->Element(XlsxWorkbookSerializerCommon::MainNs . 'cellXfs');
        if ($cellXfsContainer !== null) {
            foreach ($cellXfsContainer->Elements(XlsxWorkbookSerializerCommon::MainNs . 'xf') as $xfElement) {
                $cellXfs[] = $xfElement;
            }
        }

        if (count($cellXfs) === 0) {
            $context->CellFormats[] = StyleValue::Default()->Clone();
            $context->DefaultCellStyle = $context->CellFormats[0]->Clone();
            return $context;
        }

        for ($index = 0; $index < count($cellXfs); $index++) {
            $xf = $cellXfs[$index];
            $style = StyleValue::Default()->Clone();

            $fontId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($xf->Attribute('fontId'));
            if ($fontId !== null && $fontId >= 0 && $fontId < count($fonts)) {
                $style->Font = $fonts[$fontId]->Clone();
            }

            $fillId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($xf->Attribute('fillId'));
            if ($fillId !== null && $fillId >= 0 && $fillId < count($fills)) {
                $fill = $fills[$fillId];
                $style->Pattern = $fill->Pattern;
                $style->ForegroundColor = $fill->ForegroundColor;
                $style->BackgroundColor = $fill->BackgroundColor;
            }

            $borderId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($xf->Attribute('borderId'));
            if ($borderId !== null && $borderId >= 0 && $borderId < count($borders)) {
                $style->Borders = $borders[$borderId]->Clone();
            }

            $numFmtId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($xf->Attribute('numFmtId'));
            if ($numFmtId === null) {
                $numFmtId = 0;
            }
            $customFormat = array_key_exists($numFmtId, $customFormats) ? $customFormats[$numFmtId] : null;
            $numberFormat = new NumberFormatValue();
            $numberFormat->Number = $numFmtId;
            $numberFormat->Custom = $customFormat;
            $style->NumberFormat = $numberFormat;

            $alignment = $xf->Element(XlsxWorkbookSerializerCommon::MainNs . 'alignment');
            if ($alignment !== null) {
                $alignmentValue = new AlignmentValue();
                $alignmentValue->Horizontal = XlsxWorkbookStylesXml::ParseHorizontalAlignment((string)$alignment->Attribute('horizontal'));
                $alignmentValue->Vertical = XlsxWorkbookStylesXml::ParseVerticalAlignment((string)$alignment->Attribute('vertical'));
                $alignmentValue->WrapText = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($alignment->Attribute('wrapText'));
                $alignmentValue->IndentLevel = StyleValueSanitizer::NormalizeIndentLevel(XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('indent')));
                $alignmentValue->TextRotation = StyleValueSanitizer::NormalizeTextRotation(XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('textRotation')));
                $alignmentValue->ShrinkToFit = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($alignment->Attribute('shrinkToFit'));
                $alignmentValue->ReadingOrder = StyleValueSanitizer::NormalizeReadingOrder(XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('readingOrder')));
                $relativeIndent = XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('relativeIndent'));
                $alignmentValue->RelativeIndent = $relativeIndent !== null ? $relativeIndent : 0;
                $style->Alignment = $alignmentValue;
            }

            $protection = $xf->Element(XlsxWorkbookSerializerCommon::MainNs . 'protection');
            if ($protection !== null) {
                $locked = XlsxWorkbookStylesXml::ParseOptionalBoolAttribute($protection->Attribute('locked'));
                $protectionValue = new ProtectionValue();
                $protectionValue->IsLocked = $locked !== null ? $locked : true;
                $protectionValue->IsHidden = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($protection->Attribute('hidden'));
                $style->Protection = $protectionValue;
            }

            $context->CellFormats[] = $style;
            if (self::IsDateNumberFormat($style->NumberFormat)) {
                $context->DateStyleIndexes[$index] = true;
            }
        }

        if (count($context->CellFormats) === 0) {
            $context->CellFormats[] = StyleValue::Default()->Clone();
        }

        $context->DefaultCellStyle = $context->CellFormats[0]->Clone();
        return $context;
    }

    private static function Intern(&$indices, &$values, $value, $getKey)
    {
        $key = $getKey($value);
        if (array_key_exists($key, $indices)) {
            return $indices[$key];
        }

        $id = count($values);
        $values[] = $value;
        $indices[$key] = $id;
        return $id;
    }

    public static function GetStyleKey(StyleValue $style)
    {
        return self::GetFontKey($style->Font) . '|'
            . self::GetFillKey(self::ToFillValue($style)) . '|'
            . self::GetBordersKey($style->Borders) . '|'
            . $style->NumberFormat->Number . '|'
            . ($style->NumberFormat->Custom === null ? '' : $style->NumberFormat->Custom) . '|'
            . ((int)$style->Alignment->Horizontal) . '|'
            . ((int)$style->Alignment->Vertical) . '|'
            . ($style->Alignment->WrapText ? '1' : '0') . '|'
            . $style->Alignment->IndentLevel . '|'
            . $style->Alignment->TextRotation . '|'
            . ($style->Alignment->ShrinkToFit ? '1' : '0') . '|'
            . $style->Alignment->ReadingOrder . '|'
            . $style->Alignment->RelativeIndent . '|'
            . ($style->Protection->IsLocked ? '1' : '0') . '|'
            . ($style->Protection->IsHidden ? '1' : '0');
    }

    private static function GetFontKey(FontValue $font)
    {
        $sizeText = number_format($font->Size, 4, '.', '');
        $sizeText = rtrim(rtrim($sizeText, '0'), '.');
        if ($sizeText === '') {
            $sizeText = '0';
        }

        return $font->Name . '|'
            . $sizeText . '|'
            . ($font->Bold ? '1' : '0') . '|'
            . ($font->Italic ? '1' : '0') . '|'
            . ((int)$font->Underline) . '|'
            . ($font->StrikeThrough ? '1' : '0') . '|'
            . self::GetColorKey($font->Color);
    }

    private static function GetFillKey(FillValue $fill)
    {
        return ((int)$fill->Pattern) . '|' . self::GetColorKey($fill->ForegroundColor) . '|' . self::GetColorKey($fill->BackgroundColor);
    }

    private static function GetBordersKey(BordersValue $borders)
    {
        return self::GetBorderSideKey($borders->Left) . '|'
            . self::GetBorderSideKey($borders->Right) . '|'
            . self::GetBorderSideKey($borders->Top) . '|'
            . self::GetBorderSideKey($borders->Bottom) . '|'
            . self::GetBorderSideKey($borders->Diagonal) . '|'
            . ($borders->DiagonalUp ? '1' : '0') . '|'
            . ($borders->DiagonalDown ? '1' : '0');
    }

    private static function GetBorderSideKey(BorderSideValue $border)
    {
        return ((int)$border->Style) . '|' . self::GetColorKey($border->Color);
    }

    private static function GetColorKey(ColorValue $color)
    {
        if ($color->ThemeIndex !== null) {
            return 'T' . $color->ThemeIndex . ($color->Tint !== null ? '_' . $color->Tint : '');
        }
        if ($color->Indexed !== null) {
            return 'I' . $color->Indexed;
        }
        return self::ToArgbHex($color);
    }

    private static function ToArgbHex(ColorValue $color)
    {
        return sprintf('%02X%02X%02X%02X', $color->A, $color->R, $color->G, $color->B);
    }

    private static function GetEffectiveStyle(CellRecord $record)
    {
        if ($record->Style === null) {
            return StyleValue::Default()->Clone();
        }

        return $record->Style;
    }

    private static function ToFillValue(StyleValue $style)
    {
        $fillValue = new FillValue();
        $fillValue->Pattern = $style->Pattern;
        $fillValue->ForegroundColor = $style->ForegroundColor;
        $fillValue->BackgroundColor = $style->BackgroundColor;
        return $fillValue;
    }

    private static function ResolveNumberFormatId(NumberFormatValue $numberFormat, &$customNumberFormatIds, &$customNumberFormats, &$nextCustomNumberFormatId)
    {
        $customFormat = $numberFormat->Custom;
        if (!($customFormat === null || $customFormat === '')) {
            if (array_key_exists($customFormat, $customNumberFormatIds)) {
                return $customNumberFormatIds[$customFormat];
            }

            $customId = $nextCustomNumberFormatId;
            $nextCustomNumberFormatId++;
            $customNumberFormatIds[$customFormat] = $customId;
            $customNumberFormats[] = array($customId, $customFormat);
            return $customId;
        }

        return $numberFormat->Number >= 0 ? $numberFormat->Number : 0;
    }

    public static function StylesEqual(StyleValue $left, StyleValue $right)
    {
        return self::GetStyleKey($left) === self::GetStyleKey($right);
    }

    private static function IsDateNumberFormat(NumberFormatValue $numberFormat)
    {
        $builtIn = XlsxWorkbookSerializerCommon::builtInDateFormats();
        if (isset($builtIn[$numberFormat->Number])) {
            return true;
        }

        $customFormat = $numberFormat->Custom;
        return !($customFormat === null || $customFormat === '') && XlsxWorkbookArchiveHelpers::LooksLikeDateFormat($customFormat);
    }
}
