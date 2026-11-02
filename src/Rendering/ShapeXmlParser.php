<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class ShapeXmlParser
{
    const Xdr = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';
    const Dsp = '{http://schemas.microsoft.com/office/drawing/2008/diagram}';
    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function Parse($rawXml, $geometryFallback, $colors)
    {
        if ($rawXml === null || $rawXml === '') {
            return null;
        }

        try {
            $doc = XDocument::Parse($rawXml);
        } catch (\Throwable $exception) {
            return self::FromFallback($geometryFallback);
        }

        $root = $doc->Root;
        if ($root === null) {
            return self::FromFallback($geometryFallback);
        }

        return self::ParseShapeElement($root, self::Xdr, $geometryFallback, $colors);
    }

    // Parses a SmartArt diagram drawing part (dsp:drawing) into its laid-out shapes. Each shape
    // carries an absolute position/size (points) relative to the diagram frame.
    public static function ParseSmartArtDrawing($rawXml, $colors)
    {
        $result = array();
        if ($rawXml === null || $rawXml === '') {
            return $result;
        }

        try {
            $doc = XDocument::Parse($rawXml);
        } catch (\Throwable $exception) {
            return $result;
        }

        foreach ($doc->Descendants(self::Dsp . 'sp') as $sp) {
            $spPr = $sp->Element(self::Dsp . 'spPr');
            $xfrm = $spPr !== null ? $spPr->Element(self::A . 'xfrm') : null;
            if ($xfrm === null) {
                continue;
            }

            $off = $xfrm->Element(self::A . 'off');
            $ext = $xfrm->Element(self::A . 'ext');
            $ox = 0.0;
            $oy = 0.0;
            $cx = 0.0;
            $cy = 0.0;
            if ($off === null || $ext === null
                || !self::tryParseDouble(self::AttrOf($off, 'x'), $ox)
                || !self::tryParseDouble(self::AttrOf($off, 'y'), $oy)
                || !self::tryParseDouble(self::AttrOf($ext, 'cx'), $cx)
                || !self::tryParseDouble(self::AttrOf($ext, 'cy'), $cy)
                || $cx <= 0.0 || $cy <= 0.0) {
                continue;
            }

            $shape = self::ParseShapeElement($sp, self::Dsp, 'rect', $colors);
            $smartArtShape = new SmartArtShape();
            $smartArtShape->Shape = $shape;
            $smartArtShape->XPt = $ox / 12700.0;
            $smartArtShape->YPt = $oy / 12700.0;
            $smartArtShape->WPt = $cx / 12700.0;
            $smartArtShape->HPt = $cy / 12700.0;
            $result[] = $smartArtShape;
        }

        return $result;
    }

    private static function ParseShapeElement($root, $shapeNs, $geometryFallback, $colors)
    {
        $spPr = $root->Element($shapeNs . 'spPr');
        $style = $root->Element($shapeNs . 'style');
        $txBody = $root->Element($shapeNs . 'txBody');
        $shape = new ParsedShape();

        $xfrm = $spPr !== null ? $spPr->Element(self::A . 'xfrm') : null;
        if ($xfrm !== null) {
            $shape->FlipH = self::AttrOf($xfrm, 'flipH') === '1';
            $shape->FlipV = self::AttrOf($xfrm, 'flipV') === '1';
            $rot = 0.0;
            if (self::tryParseDouble(self::AttrOf($xfrm, 'rot'), $rot)) {
                $shape->RotationDeg = $rot / 60000.0; // DrawingML angles are 60000ths of a degree.
            }
        }

        $prstGeom = $spPr !== null ? $spPr->Element(self::A . 'prstGeom') : null;
        $prst = $prstGeom !== null ? self::AttrOf($prstGeom, 'prst') : null;
        $shape->Geometry = !($prst === null || $prst === '') ? $prst : ($geometryFallback !== null ? $geometryFallback : 'rect');
        self::ParseAdjusts($prstGeom, $shape);
        $shape->CustomGeometry = $spPr !== null ? $spPr->Element(self::A . 'custGeom') : null;

        self::ParseFill($spPr, $style, $colors, $shape);
        self::ParseLine($spPr, $style, $colors, $shape);
        self::ParseText($txBody, $style, $colors, $shape);
        $effects = $spPr !== null ? $spPr->Element(self::A . 'effectLst') : null;
        if ($effects === null && $style !== null && $colors !== null) {
            $effectIndex = 0;
            if (self::tryParseInt(self::AttrOf($style->Element(self::A . 'effectRef'), 'idx'), $effectIndex)) {
                $effects = $colors->ResolveThemeEffects($effectIndex);
            }
        }
        $shape->Reflection = $effects !== null ? $effects->Element(self::A . 'reflection') : null;
        return $shape;
    }

    private static function ParseAdjusts($prstGeom, $shape)
    {
        $avLst = $prstGeom !== null ? $prstGeom->Element(self::A . 'avLst') : null;
        if ($avLst === null) {
            return;
        }

        $values = array();
        foreach ($avLst->Elements(self::A . 'gd') as $gd) {
            $fmla = self::AttrOf($gd, 'fmla');
            $v = 0.0;
            if ($fmla !== null && $fmla !== '' && str_starts_with($fmla, 'val ')
                && self::tryParseDouble(substr($fmla, 4), $v)) {
                $values[] = $v / 100000.0; // adjust values are in 1000ths of a percent
            }
        }
        $shape->Adjusts = $values;
    }

    private static function FromFallback($geometry)
    {
        $shape = new ParsedShape();
        $shape->Geometry = $geometry !== null ? $geometry : 'rect';
        $shape->HasFill = false;
        return $shape;
    }

    private static function ParseFill($spPr, $style, $colors, $shape)
    {
        if ($spPr !== null) {
            if ($spPr->Element(self::A . 'noFill') !== null) {
                $shape->HasFill = false;
                return;
            }

            $solid = $spPr->Element(self::A . 'solidFill');
            $c = SKColors::Black();
            $pattern = $spPr->Element(self::A . 'pattFill');
            if ($pattern !== null && self::AttrOf($pattern, 'prst') === 'pct5') {
                $shape->HasFill = true;
                $shape->FillPattern = 'pct5';
                $shape->FillColor = self::TryColor(self::FillChild($pattern->Element(self::A . 'bgClr')), $colors, $c) ? $c : SKColors::White();
                $shape->PatternForeground = self::TryColor(self::FillChild($pattern->Element(self::A . 'fgClr')), $colors, $c) ? $c : SKColors::Black();
                return;
            }
            if ($solid !== null && self::TryColor(self::FillChild($solid), $colors, $c)) {
                $shape->HasFill = true;
                $shape->FillColor = $c;
                return;
            }
        }

        // Fall back to the style's fillRef (idx 0 means "no fill").
        $fillRef = $style !== null ? $style->Element(self::A . 'fillRef') : null;
        if ($fillRef !== null && self::AttrOf($fillRef, 'idx') !== '0') {
            $c = SKColors::Black();
            if (self::TryColor(self::FillChild2($fillRef), $colors, $c)) {
                $shape->HasFill = true;
                $shape->FillColor = $c;
            }
        }
    }

    private static function ParseLine($spPr, $style, $colors, $shape)
    {
        $ln = $spPr !== null ? $spPr->Element(self::A . 'ln') : null;

        if ($ln !== null && $ln->Element(self::A . 'noFill') !== null) {
            $shape->HasLine = false;
            return;
        }

        // Width
        $emu = 0.0;
        if ($ln !== null && self::tryParseDouble(self::AttrOf($ln, 'w'), $emu) && $emu > 0.0) {
            $shape->LineWidthPt = $emu / 12700.0;
        } else {
            $lnRefElement = $style !== null ? $style->Element(self::A . 'lnRef') : null;
            $lnRefIdx = $lnRefElement !== null ? self::AttrOf($lnRefElement, 'idx') : null;
            $idx = 0;
            $shape->LineWidthPt = self::tryParseInt($lnRefIdx, $idx) && $colors !== null
                ? $colors->ResolveThemeLineWidth($idx, 0.75) : 0.75;
        }

        // Color: explicit ln solidFill wins, then the style's lnRef.
        $c = SKColors::Black();
        $lnFill = $ln !== null ? $ln->Element(self::A . 'solidFill') : null;
        if ($lnFill !== null && self::TryColor(self::FillChild($lnFill), $colors, $c)) {
            $shape->HasLine = true;
            $shape->LineColor = $c;
        } else {
            $lnRef = $style !== null ? $style->Element(self::A . 'lnRef') : null;
            if ($lnRef !== null && self::TryColor(self::FillChild2($lnRef), $colors, $c)) {
                $shape->HasLine = true;
                $shape->LineColor = $c;
            }
        }

        // Arrowheads
        if ($ln !== null) {
            $head = $ln->Element(self::A . 'headEnd');
            $tail = $ln->Element(self::A . 'tailEnd');
            $shape->HeadEnd = $head !== null ? (self::AttrOf($head, 'type') ?? 'none') : null;
            $shape->TailEnd = $tail !== null ? (self::AttrOf($tail, 'type') ?? 'none') : null;
        }
    }

    private static function ParseText($txBody, $style, $colors, $shape)
    {
        if ($txBody === null) {
            return;
        }

        $sb = '';
        foreach ($txBody->Descendants(self::A . 't') as $t) {
            $sb .= $t->Value;
        }
        $shape->Text = $sb;

        // Body properties: vertical anchor and text insets.
        $bodyPr = $txBody->Element(self::A . 'bodyPr');
        if ($bodyPr !== null) {
            $anchor = self::AttrOf($bodyPr, 'anchor');
            if (!($anchor === null || $anchor === '')) {
                $shape->TextVAnchor = $anchor;
            }
            $verticalType = self::AttrOf($bodyPr, 'vert');
            if (!($verticalType === null || $verticalType === '')) {
                $shape->TextVerticalType = $verticalType;
            }
            $shape->TextWrap = self::AttrOf($bodyPr, 'wrap') === 'square';
            $ins = 0.0;
            if (self::tryParseDouble(self::AttrOf($bodyPr, 'tIns'), $ins)) {
                $shape->TextInsetTopPt = $ins / 12700.0;
            } else {
                $shape->TextInsetTopPt = 3.6;
            }
            if (self::tryParseDouble(self::AttrOf($bodyPr, 'lIns'), $ins)) {
                $shape->TextInsetLeftPt = $ins / 12700.0;
            } else {
                $shape->TextInsetLeftPt = 7.2;
            }
        }

        $pPr = self::firstOrDefaultSafe($txBody->Descendants(self::A . 'pPr'));
        $algn = $pPr !== null ? self::AttrOf($pPr, 'algn') : null;
        if (!($algn === null || $algn === '')) {
            $shape->TextHAlign = $algn;
        } else {
            $shape->TextHAlign = 'l';
        }

        // Pick up size/bold/color from the first run properties available.
        $rPr = self::firstOrDefaultSafe($txBody->Descendants(self::A . 'rPr'));
        if ($rPr === null) {
            $rPr = self::firstOrDefaultSafe($txBody->Descendants(self::A . 'endParaRPr'));
        }
        if ($rPr !== null) {
            $sz = 0.0;
            if (self::tryParseDouble(self::AttrOf($rPr, 'sz'), $sz) && $sz > 0.0) {
                $shape->TextSizePt = $sz / 100.0;
            }
            $shape->TextBold = self::AttrOf($rPr, 'b') === '1';
            if ($shape->TextBold) {
                foreach ($txBody->Descendants(self::A . 'r') as $laterRun) {
                    $laterText = $laterRun->Element(self::A . 't');
                    $laterProperties = $laterRun->Element(self::A . 'rPr');
                    if ($laterText !== null && strlen($laterText->Value) > strlen($shape->Text) / 2
                        && $laterProperties !== null && self::AttrOf($laterProperties, 'b') !== '1') {
                        $shape->TextBold = false;
                        break;
                    }
                }
            }
            $c = SKColors::Black();
            $solid = $rPr->Element(self::A . 'solidFill');
            if ($solid !== null && self::TryColor(self::FillChild($solid), $colors, $c)) {
                $shape->TextColor = $c;
                return;
            }
        }

        $fallbackColor = SKColors::Black();
        $fontRef = $style !== null ? $style->Element(self::A . 'fontRef') : null;
        if ($fontRef !== null && self::TryColor(self::FillChild2($fontRef), $colors, $fallbackColor)) {
            $shape->TextColor = $fallbackColor;
        }
    }

    private static function FillChild($solidFill)
    {
        if ($solidFill === null) {
            return null;
        }

        $child = $solidFill->Element(self::A . 'srgbClr');
        return $child !== null ? $child : $solidFill->Element(self::A . 'schemeClr');
    }

    private static function FillChild2($refEl)
    {
        // fillRef/lnRef carry a scheme or srgb color child directly.
        if ($refEl === null) {
            return null;
        }

        $child = $refEl->Element(self::A . 'srgbClr');
        return $child !== null ? $child : $refEl->Element(self::A . 'schemeClr');
    }

    private static function TryColor($colorEl, $colors, &$color)
    {
        $color = SKColors::Black();
        if ($colorEl === null) {
            return false;
        }

        if ($colorEl->Name === self::A . 'srgbClr') {
            if (self::TryHex(self::AttrOf($colorEl, 'val'), $color)) {
                $color = self::ApplyColorModifiers($colorEl, $color);
                return true;
            }

            return false;
        }

        if ($colorEl->Name === self::A . 'schemeClr') {
            $baseColor = $colors->ResolveSchemeName(self::AttrOf($colorEl, 'val'), SKColors::Gray());
            $lumMod = 0.0;
            $lumOff = 0.0;
            self::tryParseDouble(self::AttrOf($colorEl->Element(self::A . 'lumMod'), 'val'), $lumMod);
            self::tryParseDouble(self::AttrOf($colorEl->Element(self::A . 'lumOff'), 'val'), $lumOff);
            $color = ($lumMod > 0.0 || $lumOff > 0.0)
                ? RenderColor::ApplyLuma($baseColor, $lumMod > 0.0 ? $lumMod / 100000.0 : 1.0, $lumOff / 100000.0)
                : $baseColor;
            $color = self::ApplyColorModifiers($colorEl, $color);
            return true;
        }

        return false;
    }

    private static function ApplyColorModifiers($colorEl, $color)
    {
        if ($colorEl === null) {
            return $color;
        }

        $shade = self::AttrOf($colorEl->Element(self::A . 'shade'), 'val');
        $shadeValue = 0;
        if (self::tryParseInt($shade, $shadeValue)) {
            $factor = $shadeValue / 100000.0;
            $color = new SKColor(
                self::ScaleByte($color->Red, $factor),
                self::ScaleByte($color->Green, $factor),
                self::ScaleByte($color->Blue, $factor),
                $color->Alpha);
        }

        $tint = self::AttrOf($colorEl->Element(self::A . 'tint'), 'val');
        $tintValue = 0;
        if (self::tryParseInt($tint, $tintValue)) {
            $amount = $tintValue / 100000.0;
            $color = new SKColor(
                self::TintByte($color->Red, $amount),
                self::TintByte($color->Green, $amount),
                self::TintByte($color->Blue, $amount),
                $color->Alpha);
        }

        $alpha = 0;
        if (self::tryParseInt(self::AttrOf($colorEl->Element(self::A . 'alpha'), 'val'), $alpha)) {
            $color = $color->WithAlpha((int)round(max(0, min(100000, $alpha)) * 255.0 / 100000.0, 0, PHP_ROUND_HALF_EVEN));
        }

        return $color;
    }

    private static function ScaleByte($value, $factor)
    {
        // DrawingML shade operates on linear light, not encoded sRGB bytes.
        $channel = $value / 255.0;
        $linear = $channel <= 0.04045 ? $channel / 12.92 : pow(($channel + 0.055) / 1.055, 2.4);
        $linear = max(0.0, min(1.0, $linear * $factor));
        $encoded = $linear <= 0.0031308 ? $linear * 12.92 : 1.055 * pow($linear, 1.0 / 2.4) - 0.055;
        return (int)round($encoded * 255.0, 0, PHP_ROUND_HALF_EVEN);
    }

    private static function TintByte($value, $amount)
    {
        $tinted = $value + (255.0 - $value) * $amount;
        if ($tinted < 0.0) {
            $tinted = 0.0;
        }
        if ($tinted > 255.0) {
            $tinted = 255.0;
        }
        return (int)round($tinted, 0, PHP_ROUND_HALF_EVEN);
    }

    private static function TryHex($hex, &$color)
    {
        $color = SKColors::Black();
        if ($hex === null || $hex === '' || strlen($hex) < 6) {
            return false;
        }

        $offset = strlen($hex) === 8 ? 2 : 0;
        $redText = substr($hex, $offset, 2);
        $greenText = substr($hex, $offset + 2, 2);
        $blueText = substr($hex, $offset + 4, 2);
        if (ctype_xdigit($redText) && ctype_xdigit($greenText) && ctype_xdigit($blueText)) {
            $color = new SKColor(hexdec($redText), hexdec($greenText), hexdec($blueText));
            return true;
        }

        return false;
    }

    private static function AttrOf($e, $name)
    {
        if ($e === null) {
            return null;
        }
        $a = $e->Attribute($name);
        return $a !== null ? (string)$a : null;
    }

    private static function tryParseDouble($text, &$value)
    {
        if ($text === null || !is_numeric($text)) {
            return false;
        }

        $value = (float)$text;
        return true;
    }

    private static function tryParseInt($text, &$value)
    {
        if ($text === null || !is_numeric($text)) {
            return false;
        }

        $value = (int)$text;
        return true;
    }

    private static function firstOrDefaultSafe($items)
    {
        foreach ($items as $item) {
            return $item;
        }

        return null;
    }
}
