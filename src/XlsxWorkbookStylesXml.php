<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AlignmentValue;
use Aspose\Cells_FOSS\Core\BorderSideValue;
use Aspose\Cells_FOSS\Core\BordersValue;
use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Core\HorizontalAlignment;
use Aspose\Cells_FOSS\Core\NumberFormatValue;
use Aspose\Cells_FOSS\Core\ProtectionValue;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Core\VerticalAlignment;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookStylesXml
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';

    private function __construct()
    {
    }

    public static function BuildStylesheetDocument($fonts, $fills, $borders, $normalCellFormat, $cellFormats, $customNumberFormats, $differentialFormats)
    {
        $stylesheet = new XElement(self::MainNs . 'styleSheet');
        if (count($customNumberFormats) > 0) {
            $numFmtElements = array();
            for ($index = 0; $index < count($customNumberFormats); $index++) {
                $pair = $customNumberFormats[$index];
                $numFmtElements[] = new XElement(
                    self::MainNs . 'numFmt',
                    new XAttribute('numFmtId', $pair[0]),
                    new XAttribute('formatCode', $pair[1])
                );
            }
            $stylesheet->Add(new XElement(
                self::MainNs . 'numFmts',
                new XAttribute('count', count($customNumberFormats)),
                $numFmtElements
            ));
        }
        $stylesheet->Add(new XElement(
            self::MainNs . 'fonts',
            new XAttribute('count', count($fonts)),
            self::BuildFontElements($fonts)
        ));
        $stylesheet->Add(new XElement(
            self::MainNs . 'fills',
            new XAttribute('count', count($fills)),
            self::BuildFillElements($fills)
        ));
        $stylesheet->Add(new XElement(
            self::MainNs . 'borders',
            new XAttribute('count', count($borders)),
            self::BuildBorderElements($borders)
        ));
        $stylesheet->Add(new XElement(
            self::MainNs . 'cellStyleXfs',
            new XAttribute('count', 1),
            self::BuildCellStyleFormatElement($normalCellFormat)
        ));
        $stylesheet->Add(new XElement(
            self::MainNs . 'cellXfs',
            new XAttribute('count', count($cellFormats)),
            self::BuildCellFormatElements($cellFormats)
        ));
        $stylesheet->Add(new XElement(
            self::MainNs . 'cellStyles',
            new XAttribute('count', 1),
            new XElement(
                self::MainNs . 'cellStyle',
                new XAttribute('name', 'Normal'),
                new XAttribute('xfId', 0),
                new XAttribute('builtinId', 0)
            )
        ));
        if (count($differentialFormats) > 0) {
            $stylesheet->Add(new XElement(
                self::MainNs . 'dxfs',
                new XAttribute('count', count($differentialFormats)),
                self::BuildDifferentialFormatElements($differentialFormats)
            ));
        }
        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $stylesheet);
    }

    public static function ReadFontValues($root)
    {
        $fonts = array();
        $fontsElement = $root->Element(self::MainNs . 'fonts');
        $fontElements = $fontsElement === null ? array() : $fontsElement->Elements(self::MainNs . 'font');
        foreach ($fontElements as $font) {
            $fonts[] = self::ReadFontValue($font);
        }
        return $fonts;
    }

    public static function ReadFillValues($root)
    {
        $fills = array();
        $fillsElement = $root->Element(self::MainNs . 'fills');
        $fillElements = $fillsElement === null ? array() : $fillsElement->Elements(self::MainNs . 'fill');
        foreach ($fillElements as $fill) {
            $fills[] = self::ReadFillValue($fill);
        }
        return $fills;
    }

    public static function ReadBordersValues($root)
    {
        $borders = array();
        $bordersElement = $root->Element(self::MainNs . 'borders');
        $borderElements = $bordersElement === null ? array() : $bordersElement->Elements(self::MainNs . 'border');
        foreach ($borderElements as $border) {
            $borders[] = self::ReadBordersValue($border);
        }
        return $borders;
    }

    public static function ReadDifferentialStyleValues($root)
    {
        $styles = array();
        $dxfsElement = $root->Element(self::MainNs . 'dxfs');
        $dxfElements = $dxfsElement === null ? array() : $dxfsElement->Elements(self::MainNs . 'dxf');
        foreach ($dxfElements as $dxf) {
            $styles[] = self::ReadDifferentialStyleValue($dxf);
        }
        return $styles;
    }

    public static function BuildFontElements($fonts)
    {
        $elements = array();
        for ($index = 0; $index < count($fonts); $index++) {
            $elements[] = self::BuildFontElement($fonts[$index]);
        }
        return $elements;
    }

    public static function BuildFillElements($fills)
    {
        $elements = array();
        for ($index = 0; $index < count($fills); $index++) {
            $elements[] = self::BuildFillElement($fills[$index]);
        }
        return $elements;
    }

    public static function BuildBorderElements($borders)
    {
        $elements = array();
        for ($index = 0; $index < count($borders); $index++) {
            $elements[] = self::BuildBorderElement($borders[$index]);
        }
        return $elements;
    }

    public static function BuildCellFormatElements($cellFormats)
    {
        $elements = array();
        for ($index = 0; $index < count($cellFormats); $index++) {
            $elements[] = self::BuildCellFormatElement($cellFormats[$index]);
        }
        return $elements;
    }

    public static function BuildDifferentialFormatElements($differentialFormats)
    {
        $elements = array();
        for ($index = 0; $index < count($differentialFormats); $index++) {
            $elements[] = self::BuildDifferentialFormatElement($differentialFormats[$index]);
        }
        return $elements;
    }

    public static function ParseHorizontalAlignment($value)
    {
        switch ($value === null ? null : strtolower($value)) {
            case 'left':
                return HorizontalAlignment::Left;
            case 'center':
                return HorizontalAlignment::Center;
            case 'right':
                return HorizontalAlignment::Right;
            case 'fill':
                return HorizontalAlignment::Fill;
            case 'justify':
                return HorizontalAlignment::Justify;
            case 'centercontinuous':
                return HorizontalAlignment::CenterContinuous;
            case 'distributed':
                return HorizontalAlignment::Distributed;
            default:
                return HorizontalAlignment::General;
        }
    }

    public static function ParseVerticalAlignment($value)
    {
        switch ($value === null ? null : strtolower($value)) {
            case 'center':
                return VerticalAlignment::Center;
            case 'top':
                return VerticalAlignment::Top;
            case 'justify':
                return VerticalAlignment::Justify;
            case 'distributed':
                return VerticalAlignment::Distributed;
            default:
                return VerticalAlignment::Bottom;
        }
    }

    public static function ParseOptionalBoolAttribute($attribute)
    {
        if ($attribute === null) {
            return null;
        }
        return XlsxWorkbookArchiveHelpers::ParseBoolAttribute($attribute);
    }

    private static function BuildCellStyleFormatElement($cellFormat)
    {
        return self::BuildCellFormatElement($cellFormat, false);
    }

    private static function ReadDifferentialStyleValue($dxf)
    {
        $style = StyleValue::Default()->Clone();
        $font = $dxf->Element(self::MainNs . 'font');
        if ($font !== null) {
            $style->Font = self::ReadFontValue($font);
        }
        $fill = $dxf->Element(self::MainNs . 'fill');
        if ($fill !== null) {
            $fillValue = self::ReadFillValue($fill);
            $style->Pattern = $fillValue->Pattern;
            $style->ForegroundColor = $fillValue->ForegroundColor;
            $style->BackgroundColor = $fillValue->BackgroundColor;
        }
        $border = $dxf->Element(self::MainNs . 'border');
        if ($border !== null) {
            $style->Borders = self::ReadBordersValue($border);
        }
        $numFmt = $dxf->Element(self::MainNs . 'numFmt');
        if ($numFmt !== null) {
            $numberFormatValue = new NumberFormatValue();
            $numberFormatId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($numFmt->Attribute('numFmtId'));
            $numberFormatValue->Number = $numberFormatId === null ? 0 : $numberFormatId;
            $numberFormatValue->Custom = (string)$numFmt->Attribute('formatCode');
            $style->NumberFormat = $numberFormatValue;
        }
        $alignment = $dxf->Element(self::MainNs . 'alignment');
        if ($alignment !== null) {
            $alignmentValue = new AlignmentValue();
            $alignmentValue->Horizontal = self::ParseHorizontalAlignment((string)$alignment->Attribute('horizontal'));
            $alignmentValue->Vertical = self::ParseVerticalAlignment((string)$alignment->Attribute('vertical'));
            $alignmentValue->WrapText = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($alignment->Attribute('wrapText'));
            $alignmentValue->IndentLevel = StyleValueSanitizer::NormalizeIndentLevel(XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('indent')));
            $alignmentValue->TextRotation = StyleValueSanitizer::NormalizeTextRotation(XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('textRotation')));
            $alignmentValue->ShrinkToFit = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($alignment->Attribute('shrinkToFit'));
            $alignmentValue->ReadingOrder = StyleValueSanitizer::NormalizeReadingOrder(XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('readingOrder')));
            $relativeIndent = XlsxWorkbookArchiveHelpers::ParseIntAttribute($alignment->Attribute('relativeIndent'));
            $alignmentValue->RelativeIndent = $relativeIndent === null ? 0 : $relativeIndent;
            $style->Alignment = $alignmentValue;
        }
        $protection = $dxf->Element(self::MainNs . 'protection');
        if ($protection !== null) {
            $protectionValue = new ProtectionValue();
            $locked = self::ParseOptionalBoolAttribute($protection->Attribute('locked'));
            $protectionValue->IsLocked = $locked === null ? true : $locked;
            $protectionValue->IsHidden = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($protection->Attribute('hidden'));
            $style->Protection = $protectionValue;
        }
        return $style;
    }

    private static function ReadFontValue($font)
    {
        $value = new FontValue();
        $nameElement = $font->Element(self::MainNs . 'name');
        $nameAttribute = $nameElement === null ? null : $nameElement->Attribute('val');
        $value->Name = $nameAttribute === null ? 'Calibri' : (string)$nameAttribute->Value;
        $sizeElement = $font->Element(self::MainNs . 'sz');
        $sizeAttribute = $sizeElement === null ? null : $sizeElement->Attribute('val');
        $size = XlsxWorkbookArchiveHelpers::ParseDoubleAttribute($sizeAttribute);
        $value->Size = $size === null ? 11.0 : $size;
        $value->Bold = $font->Element(self::MainNs . 'b') !== null;
        $value->Italic = $font->Element(self::MainNs . 'i') !== null;
        $value->Underline = self::ParseUnderlineType($font->Element(self::MainNs . 'u'));
        $value->StrikeThrough = $font->Element(self::MainNs . 'strike') !== null;
        $value->Color = self::ReadColorValue($font->Element(self::MainNs . 'color'));
        return $value;
    }

    private static function ReadFillValue($fill)
    {
        $patternFill = $fill->Element(self::MainNs . 'patternFill');
        if ($patternFill === null) {
            return new FillValue();
        }
        $patternAttribute = $patternFill->Attribute('patternType');
        $patternType = strtolower($patternAttribute === null ? 'none' : (string)$patternAttribute->Value);
        $pattern = FillPatternKind::None;
        switch ($patternType) {
            case 'solid':
                $pattern = FillPatternKind::Solid;
                break;
            case 'mediumgray':
                $pattern = FillPatternKind::MediumGray;
                break;
            case 'darkgray':
                $pattern = FillPatternKind::DarkGray;
                break;
            case 'gray125':
                $pattern = FillPatternKind::Gray125;
                break;
            case 'gray0625':
                $pattern = FillPatternKind::Gray0625;
                break;
            case 'darkhorizontal':
                $pattern = FillPatternKind::DarkHorizontal;
                break;
            case 'darkvertical':
                $pattern = FillPatternKind::DarkVertical;
                break;
            case 'darkdown':
                $pattern = FillPatternKind::DarkDown;
                break;
            case 'darkup':
                $pattern = FillPatternKind::DarkUp;
                break;
            case 'darkgrid':
                $pattern = FillPatternKind::DarkGrid;
                break;
            case 'darktrellis':
                $pattern = FillPatternKind::DarkTrellis;
                break;
            case 'lighthorizontal':
                $pattern = FillPatternKind::LightHorizontal;
                break;
            case 'lightvertical':
                $pattern = FillPatternKind::LightVertical;
                break;
            case 'lightdown':
                $pattern = FillPatternKind::LightDown;
                break;
            case 'lightup':
                $pattern = FillPatternKind::LightUp;
                break;
            case 'lightgrid':
                $pattern = FillPatternKind::LightGrid;
                break;
            case 'lighttrellis':
                $pattern = FillPatternKind::LightTrellis;
                break;
            default:
                $pattern = FillPatternKind::None;
                break;
        }
        $value = new FillValue();
        $value->Pattern = $pattern;
        $value->ForegroundColor = self::ReadColorValue($patternFill->Element(self::MainNs . 'fgColor'));
        $value->BackgroundColor = self::ReadColorValue($patternFill->Element(self::MainNs . 'bgColor'));
        return $value;
    }

    private static function ReadBordersValue($border)
    {
        $value = new BordersValue();
        $value->Left = self::ReadBorderSideValue($border->Element(self::MainNs . 'left'));
        $value->Right = self::ReadBorderSideValue($border->Element(self::MainNs . 'right'));
        $value->Top = self::ReadBorderSideValue($border->Element(self::MainNs . 'top'));
        $value->Bottom = self::ReadBorderSideValue($border->Element(self::MainNs . 'bottom'));
        $value->Diagonal = self::ReadBorderSideValue($border->Element(self::MainNs . 'diagonal'));
        $value->DiagonalUp = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($border->Attribute('diagonalUp'));
        $value->DiagonalDown = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($border->Attribute('diagonalDown'));
        return $value;
    }

    private static function BuildFontElement($font)
    {
        $element = new XElement(self::MainNs . 'font');
        if ($font->Bold) {
            $element->Add(new XElement(self::MainNs . 'b'));
        }
        if ($font->Italic) {
            $element->Add(new XElement(self::MainNs . 'i'));
        }
        if ($font->StrikeThrough) {
            $element->Add(new XElement(self::MainNs . 'strike'));
        }
        if ($font->Underline !== FontUnderlineType::None) {
            $element->Add(self::BuildUnderlineElement($font->Underline));
        }
        $sizeText = number_format($font->Size, 4, '.', '');
        $sizeText = rtrim(rtrim($sizeText, '0'), '.');
        if ($sizeText === '') {
            $sizeText = '0';
        }
        $element->Add(new XElement(self::MainNs . 'sz', new XAttribute('val', $sizeText)));
        $colorElement = self::BuildColorElement('color', $font->Color);
        if ($colorElement !== null) {
            $element->Add($colorElement);
        }
        $element->Add(new XElement(self::MainNs . 'name', new XAttribute('val', $font->Name)));
        return $element;
    }

    private static function ParseUnderlineType($underlineElement)
    {
        if ($underlineElement === null) {
            return FontUnderlineType::None;
        }

        $value = (string)$underlineElement->Attribute('val');
        if ($value === null || $value === '') {
            return FontUnderlineType::Single;
        }

        switch ($value) {
            case 'single':
                return FontUnderlineType::Single;
            case 'double':
                return FontUnderlineType::Double;
            case 'singleAccounting':
                return FontUnderlineType::Accounting;
            case 'doubleAccounting':
                return FontUnderlineType::DoubleAccounting;
            default:
                return FontUnderlineType::Single;
        }
    }

    private static function BuildUnderlineElement($underlineType)
    {
        switch ($underlineType) {
            case FontUnderlineType::Double:
                return new XElement(self::MainNs . 'u', new XAttribute('val', 'double'));
            case FontUnderlineType::Accounting:
                return new XElement(self::MainNs . 'u', new XAttribute('val', 'singleAccounting'));
            case FontUnderlineType::DoubleAccounting:
                return new XElement(self::MainNs . 'u', new XAttribute('val', 'doubleAccounting'));
            case FontUnderlineType::Single:
                return new XElement(self::MainNs . 'u');
            default:
                return new XElement(self::MainNs . 'u');
        }
    }

    private static function BuildFillElement($fill)
    {
        $patternFill = new XElement(self::MainNs . 'patternFill');
        switch ($fill->Pattern) {
            case FillPatternKind::Solid:
                $patternFill->SetAttributeValue('patternType', 'solid');
                break;
            case FillPatternKind::MediumGray:
                $patternFill->SetAttributeValue('patternType', 'mediumGray');
                break;
            case FillPatternKind::DarkGray:
                $patternFill->SetAttributeValue('patternType', 'darkGray');
                break;
            case FillPatternKind::Gray125:
                $patternFill->SetAttributeValue('patternType', 'gray125');
                break;
            case FillPatternKind::Gray0625:
                $patternFill->SetAttributeValue('patternType', 'gray0625');
                break;
            case FillPatternKind::DarkHorizontal:
                $patternFill->SetAttributeValue('patternType', 'darkHorizontal');
                break;
            case FillPatternKind::DarkVertical:
                $patternFill->SetAttributeValue('patternType', 'darkVertical');
                break;
            case FillPatternKind::DarkDown:
                $patternFill->SetAttributeValue('patternType', 'darkDown');
                break;
            case FillPatternKind::DarkUp:
                $patternFill->SetAttributeValue('patternType', 'darkUp');
                break;
            case FillPatternKind::DarkGrid:
                $patternFill->SetAttributeValue('patternType', 'darkGrid');
                break;
            case FillPatternKind::DarkTrellis:
                $patternFill->SetAttributeValue('patternType', 'darkTrellis');
                break;
            case FillPatternKind::LightHorizontal:
                $patternFill->SetAttributeValue('patternType', 'lightHorizontal');
                break;
            case FillPatternKind::LightVertical:
                $patternFill->SetAttributeValue('patternType', 'lightVertical');
                break;
            case FillPatternKind::LightDown:
                $patternFill->SetAttributeValue('patternType', 'lightDown');
                break;
            case FillPatternKind::LightUp:
                $patternFill->SetAttributeValue('patternType', 'lightUp');
                break;
            case FillPatternKind::LightGrid:
                $patternFill->SetAttributeValue('patternType', 'lightGrid');
                break;
            case FillPatternKind::LightTrellis:
                $patternFill->SetAttributeValue('patternType', 'lightTrellis');
                break;
            default:
                $patternFill->SetAttributeValue('patternType', 'none');
                break;
        }
        $foregroundColor = self::BuildColorElement('fgColor', $fill->ForegroundColor);
        if ($foregroundColor !== null) {
            $patternFill->Add($foregroundColor);
        }
        $backgroundColor = self::BuildColorElement('bgColor', $fill->BackgroundColor);
        if ($backgroundColor !== null) {
            $patternFill->Add($backgroundColor);
        }
        return new XElement(self::MainNs . 'fill', $patternFill);
    }

    private static function BuildBorderElement($borders)
    {
        $element = new XElement(
            self::MainNs . 'border',
            self::BuildBorderSideElement('left', $borders->Left),
            self::BuildBorderSideElement('right', $borders->Right),
            self::BuildBorderSideElement('top', $borders->Top),
            self::BuildBorderSideElement('bottom', $borders->Bottom),
            self::BuildBorderSideElement('diagonal', $borders->Diagonal)
        );
        if ($borders->DiagonalUp) {
            $element->SetAttributeValue('diagonalUp', 1);
        }
        if ($borders->DiagonalDown) {
            $element->SetAttributeValue('diagonalDown', 1);
        }
        return $element;
    }

    private static function BuildCellFormatElement($cellFormat, $includeXfId = true)
    {
        $element = new XElement(
            self::MainNs . 'xf',
            new XAttribute('numFmtId', $cellFormat->NumFmtId),
            new XAttribute('fontId', $cellFormat->FontId),
            new XAttribute('fillId', $cellFormat->FillId),
            new XAttribute('borderId', $cellFormat->BorderId)
        );
        if ($includeXfId) {
            $element->SetAttributeValue('xfId', 0);
        }
        if ($cellFormat->NumFmtId != 0) {
            $element->SetAttributeValue('applyNumberFormat', 1);
        }
        if ($cellFormat->FontId != 0) {
            $element->SetAttributeValue('applyFont', 1);
        }
        if ($cellFormat->FillId != 0) {
            $element->SetAttributeValue('applyFill', 1);
        }
        if ($cellFormat->BorderId != 0) {
            $element->SetAttributeValue('applyBorder', 1);
        }
        $alignmentElement = self::BuildAlignmentElement($cellFormat->Alignment);
        if ($alignmentElement !== null) {
            $element->SetAttributeValue('applyAlignment', 1);
            $element->Add($alignmentElement);
        }
        $protectionElement = self::BuildProtectionElement($cellFormat->Protection);
        if ($protectionElement !== null) {
            $element->SetAttributeValue('applyProtection', 1);
            $element->Add($protectionElement);
        }
        return $element;
    }

    private static function BuildDifferentialFormatElement($style)
    {
        $element = new XElement(self::MainNs . 'dxf');
        if (!XlsxWorkbookStylesValueHelpers::FontEquals($style->Font, StyleValue::Default()->Font)) {
            $element->Add(self::BuildFontElement($style->Font));
        }
        if ($style->Pattern !== FillPatternKind::None || !XlsxWorkbookStylesValueHelpers::IsEmptyColor($style->ForegroundColor) || !XlsxWorkbookStylesValueHelpers::IsEmptyColor($style->BackgroundColor)) {
            $fillValue = new FillValue();
            $fillValue->Pattern = $style->Pattern;
            $fillValue->ForegroundColor = $style->ForegroundColor;
            $fillValue->BackgroundColor = $style->BackgroundColor;
            $element->Add(self::BuildFillElement($fillValue));
        }
        if (!XlsxWorkbookStylesValueHelpers::BordersEqual($style->Borders, StyleValue::Default()->Borders)) {
            $element->Add(self::BuildBorderElement($style->Borders));
        }
        if ($style->NumberFormat->Number != 0 || !($style->NumberFormat->Custom === null || $style->NumberFormat->Custom === '')) {
            $numFmtElement = new XElement(
                self::MainNs . 'numFmt',
                new XAttribute('numFmtId', $style->NumberFormat->Number >= 0 ? $style->NumberFormat->Number : 0)
            );
            if (!($style->NumberFormat->Custom === null || $style->NumberFormat->Custom === '')) {
                $numFmtElement->SetAttributeValue('formatCode', $style->NumberFormat->Custom);
            }
            $element->Add($numFmtElement);
        }
        $alignmentElement = self::BuildAlignmentElement($style->Alignment);
        if ($alignmentElement !== null) {
            $element->Add($alignmentElement);
        }
        $protectionElement = self::BuildProtectionElement($style->Protection);
        if ($protectionElement !== null) {
            $element->Add($protectionElement);
        }
        return $element;
    }

    private static function ReadBorderSideValue($side)
    {
        if ($side === null) {
            return new BorderSideValue();
        }
        $value = new BorderSideValue();
        $value->Style = XlsxWorkbookStylesValueHelpers::ParseBorderStyle((string)$side->Attribute('style'));
        $value->Color = self::ReadColorValue($side->Element(self::MainNs . 'color'));
        return $value;
    }

    private static function ReadColorValue($colorElement)
    {
        if ($colorElement === null) {
            return new ColorValue();
        }

        $themeAttr = (string)$colorElement->Attribute('theme');
        if (!($themeAttr === null || $themeAttr === '')) {
            if (is_numeric($themeAttr)) {
                $themeIndex = intval($themeAttr);
                $tint = null;
                $tintAttr = (string)$colorElement->Attribute('tint');
                if (!($tintAttr === null || $tintAttr === '')) {
                    if (is_numeric($tintAttr)) {
                        $tint = floatval($tintAttr);
                    }
                }
                return new ColorValue($themeIndex, $tint);
            }
        }

        $indexedAttr = (string)$colorElement->Attribute('indexed');
        if (!($indexedAttr === null || $indexedAttr === '')) {
            if (is_numeric($indexedAttr)) {
                return new ColorValue(intval($indexedAttr));
            }
        }

        $rgb = (string)$colorElement->Attribute('rgb');
        if ($rgb === null || trim($rgb) === '') {
            return new ColorValue();
        }
        $rgbValue = strtoupper(trim($rgb));
        if (strlen($rgbValue) === 6) {
            $rgbValue = 'FF' . $rgbValue;
        }
        if (strlen($rgbValue) !== 8) {
            return new ColorValue();
        }
        if (!ctype_xdigit(substr($rgbValue, 0, 2))) {
            return new ColorValue();
        }
        if (!ctype_xdigit(substr($rgbValue, 2, 2))) {
            return new ColorValue();
        }
        if (!ctype_xdigit(substr($rgbValue, 4, 2))) {
            return new ColorValue();
        }
        if (!ctype_xdigit(substr($rgbValue, 6, 2))) {
            return new ColorValue();
        }
        $a = hexdec(substr($rgbValue, 0, 2));
        $r = hexdec(substr($rgbValue, 2, 2));
        $g = hexdec(substr($rgbValue, 4, 2));
        $b = hexdec(substr($rgbValue, 6, 2));
        return new ColorValue($a, $r, $g, $b);
    }

    private static function BuildColorElement($elementName, $color)
    {
        if (XlsxWorkbookStylesValueHelpers::IsEmptyColor($color)) {
            return null;
        }
        if ($color->ThemeIndex !== null) {
            $element = new XElement(
                self::MainNs . $elementName,
                new XAttribute('theme', (string)$color->ThemeIndex)
            );
            if ($color->Tint !== null) {
                $element->SetAttributeValue('tint', (string)$color->Tint);
            }
            return $element;
        }
        if ($color->Indexed !== null) {
            return new XElement(
                self::MainNs . $elementName,
                new XAttribute('indexed', (string)$color->Indexed)
            );
        }
        return new XElement(self::MainNs . $elementName, new XAttribute('rgb', XlsxWorkbookStylesValueHelpers::ToArgbHex($color)));
    }

    private static function BuildBorderSideElement($sideName, $side)
    {
        $element = new XElement(self::MainNs . $sideName);
        $styleName = XlsxWorkbookStylesValueHelpers::GetBorderStyleName($side->Style);
        if (!($styleName === null || $styleName === '')) {
            $element->SetAttributeValue('style', $styleName);
        }
        $colorElement = self::BuildColorElement('color', $side->Color);
        if ($colorElement !== null) {
            $element->Add($colorElement);
        }
        return $element;
    }

    private static function BuildAlignmentElement($alignment)
    {
        $element = new XElement(self::MainNs . 'alignment');
        $hasValue = false;
        $horizontal = XlsxWorkbookStylesValueHelpers::GetHorizontalAlignmentName($alignment->Horizontal);
        if (!($horizontal === null || $horizontal === '')) {
            $element->SetAttributeValue('horizontal', $horizontal);
            $hasValue = true;
        }
        $vertical = XlsxWorkbookStylesValueHelpers::GetVerticalAlignmentName($alignment->Vertical);
        if (!($vertical === null || $vertical === '')) {
            $element->SetAttributeValue('vertical', $vertical);
            $hasValue = true;
        }
        if ($alignment->WrapText) {
            $element->SetAttributeValue('wrapText', 1);
            $hasValue = true;
        }
        if ($alignment->IndentLevel > 0) {
            $element->SetAttributeValue('indent', $alignment->IndentLevel);
            $hasValue = true;
        }
        if ($alignment->TextRotation != 0) {
            $element->SetAttributeValue('textRotation', $alignment->TextRotation);
            $hasValue = true;
        }
        if ($alignment->ShrinkToFit) {
            $element->SetAttributeValue('shrinkToFit', 1);
            $hasValue = true;
        }
        if ($alignment->ReadingOrder != 0) {
            $element->SetAttributeValue('readingOrder', $alignment->ReadingOrder);
            $hasValue = true;
        }
        if ($alignment->RelativeIndent != 0) {
            $element->SetAttributeValue('relativeIndent', $alignment->RelativeIndent);
            $hasValue = true;
        }
        return $hasValue ? $element : null;
    }

    private static function BuildProtectionElement($protection)
    {
        $element = new XElement(self::MainNs . 'protection');
        $hasValue = false;
        if (!$protection->IsLocked) {
            $element->SetAttributeValue('locked', 0);
            $hasValue = true;
        }
        if ($protection->IsHidden) {
            $element->SetAttributeValue('hidden', 1);
            $hasValue = true;
        }
        return $hasValue ? $element : null;
    }
}
