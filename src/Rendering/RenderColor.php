<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class RenderColor
{
    private static $DefaultScheme = null;

    private static $IndexedPalette = null;

    private static $DefaultThemeLineWidthsPt = array(0.75, 0.5, 1.0, 1.5);

    private $scheme;

    private $themeLineWidthsPt;

    private $themeEffects = array();

    public $MajorFontName;

    public $MinorFontName;

    private function __construct($scheme, $themeLineWidthsPt, $majorFontName, $minorFontName)
    {
        $this->scheme = $scheme;
        $this->themeLineWidthsPt = $themeLineWidthsPt;
        $this->MajorFontName = $majorFontName;
        $this->MinorFontName = $minorFontName;
    }

    private static function defaultScheme()
    {
        if (self::$DefaultScheme === null) {
            self::$DefaultScheme = array(
                new SKColor(0x00, 0x00, 0x00),
                new SKColor(0xFF, 0xFF, 0xFF),
                new SKColor(0x44, 0x54, 0x6A),
                new SKColor(0xE7, 0xE6, 0xE6),
                new SKColor(0x44, 0x72, 0xC4),
                new SKColor(0xED, 0x7D, 0x31),
                new SKColor(0xA5, 0xA5, 0xA5),
                new SKColor(0xFF, 0xC0, 0x00),
                new SKColor(0x5B, 0x9B, 0xD5),
                new SKColor(0x70, 0xAD, 0x47),
                new SKColor(0x05, 0x63, 0xC1),
                new SKColor(0x95, 0x4F, 0x72),
            );
        }

        return self::$DefaultScheme;
    }

    private static function indexedPalette()
    {
        if (self::$IndexedPalette === null) {
            self::$IndexedPalette = self::BuildIndexedPalette();
        }

        return self::$IndexedPalette;
    }

    public static function FromWorkbook($model)
    {
        $themeXml = $model !== null ? $model->RawThemeXml : null;
        $scheme = self::ParseThemeScheme($themeXml);
        $lineWidths = self::ParseThemeLineWidths($themeXml);
        $result = new RenderColor($scheme !== null ? $scheme : self::defaultScheme(),
            $lineWidths !== null ? $lineWidths : self::$DefaultThemeLineWidthsPt,
            self::ParseThemeFontName($themeXml, true) !== null ? self::ParseThemeFontName($themeXml, true) : 'Calibri Light',
            self::ParseThemeFontName($themeXml, false) !== null ? self::ParseThemeFontName($themeXml, false) : 'Calibri');
        if ($themeXml !== null && $themeXml !== '') {
            try {
                $a = '{http://schemas.openxmlformats.org/drawingml/2006/main}';
                $doc = XDocument::Parse($themeXml);
                $styles = self::firstOrDefaultSafe($doc->Descendants($a . 'effectStyleLst'));
                if ($styles !== null) {
                    foreach ($styles->Elements($a . 'effectStyle') as $style) {
                        $result->themeEffects[] = $style->Element($a . 'effectLst');
                    }
                }
            } catch (\Throwable $exception) {
            }
        }
        return $result;
    }

    public function ResolveThemeEffects($index)
    {
        return $index > 0 && $index <= count($this->themeEffects) ? $this->themeEffects[$index - 1] : null;
    }

    private static function ParseThemeFontName($rawThemeXml, $major)
    {
        if ($rawThemeXml === null || trim($rawThemeXml) === '') {
            return null;
        }

        try {
            $a = '{http://schemas.openxmlformats.org/drawingml/2006/main}';
            $doc = XDocument::Parse($rawThemeXml);
            $fontScheme = self::firstOrDefaultSafe($doc->Descendants($a . 'fontScheme'));
            $fontGroup = $fontScheme !== null ? $fontScheme->Element($a . ($major ? 'majorFont' : 'minorFont')) : null;
            $latin = $fontGroup !== null ? $fontGroup->Element($a . 'latin') : null;
            $attribute = $latin !== null ? $latin->Attribute('typeface') : null;
            $name = $attribute !== null ? (string)$attribute : null;
            return $name === null || trim($name) === '' ? null : $name;
        } catch (\Throwable $exception) {
            return null;
        }
    }

    public function ResolveThemeLineWidth($index, $fallback)
    {
        return $index >= 0 && $index < count($this->themeLineWidthsPt) ? $this->themeLineWidthsPt[$index] : $fallback;
    }

    public function Resolve($color, $fallback)
    {
        if ($color->ThemeIndex !== null) {
            $baseColor = $this->ResolveTheme($color->ThemeIndex);
            return self::ApplyTint($baseColor, $color->Tint !== null ? $color->Tint : 0.0);
        }

        if ($color->Indexed !== null) {
            return self::ResolveIndexed($color->Indexed, $fallback);
        }

        if ($color->A === 0 && $color->R === 0 && $color->G === 0 && $color->B === 0) {
            return $fallback;
        }

        return new SKColor($color->R, $color->G, $color->B, $color->A);
    }

    public function ResolveSchemeName($name, $fallback)
    {
        switch ($name) {
            case 'dk1':
            case 'tx1':
                $index = 0;
                break;
            case 'lt1':
            case 'bg1':
                $index = 1;
                break;
            case 'dk2':
            case 'tx2':
                $index = 2;
                break;
            case 'lt2':
            case 'bg2':
                $index = 3;
                break;
            case 'accent1':
                $index = 4;
                break;
            case 'accent2':
                $index = 5;
                break;
            case 'accent3':
                $index = 6;
                break;
            case 'accent4':
                $index = 7;
                break;
            case 'accent5':
                $index = 8;
                break;
            case 'accent6':
                $index = 9;
                break;
            case 'hlink':
                $index = 10;
                break;
            case 'folHlink':
                $index = 11;
                break;
            default:
                return $fallback;
        }

        return $index >= 0 && $index < count($this->scheme) ? $this->scheme[$index] : $fallback;
    }

    public static function ApplyLuma($color, $lumMod, $lumOff)
    {
        self::ToHsl($color, $hue, $saturation, $luminance);
        $luminance = self::ClampUnit($luminance * $lumMod + $lumOff);
        return self::FromHsl($hue, $saturation, $luminance, $color->Alpha);
    }

    private static function ToHsl($color, &$hue, &$saturation, &$luminance)
    {
        $red = $color->Red / 255.0;
        $green = $color->Green / 255.0;
        $blue = $color->Blue / 255.0;
        $maximum = max($red, max($green, $blue));
        $minimum = min($red, min($green, $blue));
        $luminance = ($maximum + $minimum) * 0.5;
        if (abs($maximum - $minimum) < 0.0000001) {
            $hue = 0.0;
            $saturation = 0.0;
            return;
        }

        $delta = $maximum - $minimum;
        $saturation = $luminance > 0.5
            ? $delta / (2.0 - $maximum - $minimum)
            : $delta / ($maximum + $minimum);
        if (abs($maximum - $red) < 0.0000001) {
            $hue = ($green - $blue) / $delta + ($green < $blue ? 6.0 : 0.0);
        } elseif (abs($maximum - $green) < 0.0000001) {
            $hue = ($blue - $red) / $delta + 2.0;
        } else {
            $hue = ($red - $green) / $delta + 4.0;
        }

        $hue /= 6.0;
    }

    private static function FromHsl($hue, $saturation, $luminance, $alpha)
    {
        if ($saturation <= 0.0) {
            $gray = self::ToColorByte($luminance);
            return new SKColor($gray, $gray, $gray, $alpha);
        }

        $q = $luminance < 0.5
            ? $luminance * (1.0 + $saturation)
            : $luminance + $saturation - $luminance * $saturation;
        $p = 2.0 * $luminance - $q;
        return new SKColor(
            self::ToColorByte(self::HueToRgb($p, $q, $hue + 1.0 / 3.0)),
            self::ToColorByte(self::HueToRgb($p, $q, $hue)),
            self::ToColorByte(self::HueToRgb($p, $q, $hue - 1.0 / 3.0)),
            $alpha);
    }

    private static function HueToRgb($p, $q, $hue)
    {
        if ($hue < 0.0) {
            $hue += 1.0;
        }
        if ($hue > 1.0) {
            $hue -= 1.0;
        }
        if ($hue < 1.0 / 6.0) {
            return $p + ($q - $p) * 6.0 * $hue;
        }
        if ($hue < 0.5) {
            return $q;
        }
        if ($hue < 2.0 / 3.0) {
            return $p + ($q - $p) * (2.0 / 3.0 - $hue) * 6.0;
        }
        return $p;
    }

    private static function ClampUnit($value)
    {
        if ($value < 0.0) {
            return 0.0;
        }
        if ($value > 1.0) {
            return 1.0;
        }
        return $value;
    }

    private static function ToColorByte($value)
    {
        return (int)round(self::ClampUnit($value) * 255.0, 0, PHP_ROUND_HALF_EVEN);
    }

    private function ResolveTheme($themeIndex)
    {
        switch ($themeIndex) {
            case 0:
                $schemeIndex = 1;
                break;
            case 1:
                $schemeIndex = 0;
                break;
            case 2:
                $schemeIndex = 3;
                break;
            case 3:
                $schemeIndex = 2;
                break;
            default:
                $schemeIndex = $themeIndex;
                break;
        }

        if ($schemeIndex >= 0 && $schemeIndex < count($this->scheme)) {
            return $this->scheme[$schemeIndex];
        }

        return SKColors::Black();
    }

    private static function ResolveIndexed($index, $fallback)
    {
        if ($index === 64 || $index === 65) {
            return $fallback;
        }

        $palette = self::indexedPalette();
        if ($index >= 0 && $index < count($palette)) {
            return $palette[$index];
        }

        return $fallback;
    }

    private static function ApplyTint($color, $tint)
    {
        if ($tint === 0.0) {
            return $color;
        }

        return new SKColor(self::TintChannel($color->Red, $tint), self::TintChannel($color->Green, $tint), self::TintChannel($color->Blue, $tint), $color->Alpha);
    }

    private static function TintChannel($channel, $tint)
    {
        if ($tint < 0.0) {
            $value = $channel * (1.0 + $tint);
        } else {
            $value = $channel * (1.0 - $tint) + 255.0 * $tint;
        }

        if ($value < 0.0) {
            $value = 0.0;
        }
        if ($value > 255.0) {
            $value = 255.0;
        }
        return (int)round($value, 0, PHP_ROUND_HALF_EVEN);
    }

    private static function ParseThemeLineWidths($rawThemeXml)
    {
        if ($rawThemeXml === null || trim($rawThemeXml) === '') {
            return null;
        }

        try {
            $a = '{http://schemas.openxmlformats.org/drawingml/2006/main}';
            $doc = XDocument::Parse($rawThemeXml);
            $lineStyles = self::firstOrDefaultSafe($doc->Descendants($a . 'lnStyleLst'));
            if ($lineStyles === null) {
                return null;
            }

            $widths = array();
            $widths[] = self::$DefaultThemeLineWidthsPt[0];
            foreach ($lineStyles->Elements($a . 'ln') as $line) {
                $fallback = count($widths) < count(self::$DefaultThemeLineWidthsPt)
                    ? self::$DefaultThemeLineWidthsPt[count($widths)] : 0.75;
                $attribute = $line->Attribute('w');
                $emu = $attribute !== null ? (string)$attribute : '';
                $widths[] = is_numeric($emu) && (float)$emu > 0.0 ? (float)$emu / 12700.0 : $fallback;
            }

            return count($widths) > 1 ? $widths : null;
        } catch (\Throwable $exception) {
            return null;
        }
    }

    private static function ParseThemeScheme($rawThemeXml)
    {
        if ($rawThemeXml === null || trim($rawThemeXml) === '') {
            return null;
        }

        try {
            $a = '{http://schemas.openxmlformats.org/drawingml/2006/main}';
            $doc = XDocument::Parse($rawThemeXml);
            $clrScheme = self::firstOrDefaultSafe($doc->Descendants($a . 'clrScheme'));
            if ($clrScheme === null) {
                return null;
            }

            $order = array('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink');
            $result = array();
            $defaults = self::defaultScheme();
            for ($i = 0; $i < count($order); $i++) {
                $element = $clrScheme->Element($a . $order[$i]);
                if ($element !== null && self::TryParseSchemeColor($element, $a, $parsed)) {
                    $result[$i] = $parsed;
                } else {
                    $result[$i] = $defaults[$i];
                }
            }

            return $result;
        } catch (\Throwable $exception) {
            return null;
        }
    }

    private static function TryParseSchemeColor($colorSlot, $a, &$color)
    {
        $color = SKColors::Black();

        $srgb = $colorSlot->Element($a . 'srgbClr');
        if ($srgb !== null) {
            $attribute = $srgb->Attribute('val');
            $val = $attribute !== null ? (string)$attribute : '';
            return self::TryParseHex($val, $color);
        }

        $sys = $colorSlot->Element($a . 'sysClr');
        if ($sys !== null) {
            $attribute = $sys->Attribute('lastClr');
            $last = $attribute !== null ? (string)$attribute : '';
            if (self::TryParseHex($last, $color)) {
                return true;
            }

            $attribute = $sys->Attribute('val');
            $name = $attribute !== null ? (string)$attribute : '';
            if (strcasecmp($name, 'window') === 0) {
                $color = SKColors::White();
                return true;
            }

            if (strcasecmp($name, 'windowText') === 0) {
                $color = SKColors::Black();
                return true;
            }
        }

        return false;
    }

    private static function TryParseHex($hex, &$color)
    {
        $color = SKColors::Black();
        if ($hex === null || $hex === '' || strlen($hex) < 6) {
            return false;
        }

        $offset = strlen($hex) === 8 ? 2 : 0;
        if (self::TryHexByte($hex, $offset, $r) && self::TryHexByte($hex, $offset + 2, $g) && self::TryHexByte($hex, $offset + 4, $b)) {
            $color = new SKColor($r, $g, $b);
            return true;
        }

        return false;
    }

    private static function TryHexByte($s, $index, &$value)
    {
        $value = 0;
        if ($index + 2 > strlen($s)) {
            return false;
        }

        $text = substr($s, $index, 2);
        if (ctype_xdigit($text)) {
            $value = hexdec($text);
            return true;
        }

        return false;
    }

    private static function BuildIndexedPalette()
    {
        $rgb = array(
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
            0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
            0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
            0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
            0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
            0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
            0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
        );

        $palette = array();
        for ($i = 0; $i < count($rgb); $i++) {
            $palette[$i] = new SKColor(($rgb[$i] >> 16) & 0xFF, ($rgb[$i] >> 8) & 0xFF, $rgb[$i] & 0xFF);
        }

        return $palette;
    }

    private static function firstOrDefaultSafe($items)
    {
        foreach ($items as $item) {
            return $item;
        }

        return null;
    }
}
