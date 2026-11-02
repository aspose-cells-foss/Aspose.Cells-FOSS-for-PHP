<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKFontHinting;
use Aspose\Cells_FOSS\Skia\SKFontManager;
use Aspose\Cells_FOSS\Skia\SKFontStyleSlant;
use Aspose\Cells_FOSS\Skia\SKFontStyleWeight;
use Aspose\Cells_FOSS\Skia\SKFontStyleWidth;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKTypeface;

final class FontRegistry
{
    private $cache = array();

    private $fallbacks = array();

    private $fallbacksInitialized = false;

    public $DefaultFontName;

    private static $ExcelDigitWidthPxAt11 = array(
        'SimSun' => 7.34,
        '宋体' => 7.34,
        'NSimSun' => 7.34,
        '新宋体' => 7.34,
    );

    public function Resolve($font)
    {
        $name = $font !== null && $font->Name !== null && $font->Name !== '' ? $font->Name : 'Calibri';
        $weight = ($font !== null && $font->Bold) ? SKFontStyleWeight::Bold : SKFontStyleWeight::Normal;
        if ($font !== null && !$font->Bold && $name !== '' && preg_match('/ Light$/i', $name)) {
            $weight = SKFontStyleWeight::Light;
        }
        $slant = ($font !== null && $font->Italic) ? SKFontStyleSlant::Italic : SKFontStyleSlant::Upright;
        $key = $name . '|' . $weight . '|' . $slant;

        if (array_key_exists($key, $this->cache)) {
            return $this->cache[$key];
        }

        $typeface = SKTypeface::FromFamilyName($name, $weight, SKFontStyleWidth::Normal, $slant);
        if ($typeface === null) {
            $typeface = SKTypeface::Default();
        }
        if (preg_match('/^Aptos/i', $name)
            && !($typeface->FamilyName !== null && strcasecmp($typeface->FamilyName, $name) === 0)) {
            $substitute = preg_match('/Display$/i', $name) ? 'Calibri Light' : 'Calibri';
            $substituteFace = SKTypeface::FromFamilyName($substitute, $weight, SKFontStyleWidth::Normal, $slant);
            if ($substituteFace !== null) {
                $typeface = $substituteFace;
            }
        }
        $this->cache[$key] = $typeface;
        return $typeface;
    }

    public function ResolveForCodepoint($primary, $codepoint)
    {
        if ($primary !== null && $primary->ContainsGlyph($codepoint)) {
            return $primary;
        }

        $this->EnsureFallbacks();
        for ($i = 0; $i < count($this->fallbacks); $i++) {
            if ($this->fallbacks[$i] !== null && $this->fallbacks[$i]->ContainsGlyph($codepoint)) {
                return $this->fallbacks[$i];
            }
        }

        $matched = SKFontManager::DefaultManager()->MatchCharacter($codepoint);
        if ($matched !== null) {
            $this->fallbacks[] = $matched;
            return $matched;
        }

        return $primary !== null ? $primary : SKTypeface::Default();
    }

    private function EnsureFallbacks()
    {
        if ($this->fallbacksInitialized) {
            return;
        }

        $this->fallbacksInitialized = true;

        if ($this->DefaultFontName !== null && $this->DefaultFontName !== '') {
            $preferred = SKTypeface::FromFamilyName($this->DefaultFontName);
            if ($preferred !== null) {
                $this->fallbacks[] = $preferred;
            }
        }

        $candidates = array(
            'Noto Sans CJK SC', 'Noto Sans CJK TC', 'Noto Sans CJK JP', 'Noto Sans',
            'Microsoft YaHei', 'SimSun', 'PingFang SC', 'Hiragino Sans', 'Arial Unicode MS',
        );

        foreach ($candidates as $candidate) {
            $typeface = SKTypeface::FromFamilyName($candidate);
            if ($typeface !== null && !($typeface->FamilyName !== null && $typeface->FamilyName === SKTypeface::Default()->FamilyName)) {
                $this->fallbacks[] = $typeface;
            }
        }
    }

    public function MeasureMaxDigitWidth($normalFont)
    {
        $typeface = $this->Resolve($normalFont);
        $sizePt = $normalFont !== null ? $normalFont->Size : 11.0;

        if ($this->TryGetDigitWidthOverride($normalFont, $typeface, $overridePxAt11)) {
            return $overridePxAt11 * ($sizePt / 11.0);
        }

        $sizePx = $sizePt * RenderUnits::PointsToPixels;
        $font = new SKFont($typeface, $sizePx);
        try {
            $font->Subpixel = true;
            $font->Hinting = SKFontHinting::None;
            $font->LinearMetrics = true;
            $paint = new SKPaint($font);
            try {
                $width = $paint->MeasureText('0');
                if ($width <= 0.0) {
                    return 7.5;
                }

                if ($typeface->FamilyName !== null && strcasecmp($typeface->FamilyName, 'Tahoma') === 0) {
                    return $width;
                }

                return ceil($width * 2.0) / 2.0;
            } finally {
                $paint->Dispose();
            }
        } finally {
            $font->Dispose();
        }
    }

    private function TryGetDigitWidthOverride($font, $typeface, &$pxAt11)
    {
        if ($font !== null && $font->Name !== null && $font->Name !== '') {
            if (strcasecmp($font->Name, 'DengXian') === 0 || $font->Name === '等线') {
                $pxAt11 = 7.75;
                return true;
            }
        }

        if ($typeface !== null && $typeface->FamilyName !== null && $typeface->FamilyName !== '') {
            if (strcasecmp($typeface->FamilyName, 'DengXian') === 0 || $typeface->FamilyName === '等线') {
                $pxAt11 = 7.75;
                return true;
            }
        }

        if ($font !== null && $font->Name !== null && $font->Name !== '' && array_key_exists($font->Name, self::$ExcelDigitWidthPxAt11)) {
            $pxAt11 = self::$ExcelDigitWidthPxAt11[$font->Name];
            return true;
        }

        if ($typeface !== null && $typeface->FamilyName !== null && $typeface->FamilyName !== ''
            && array_key_exists($typeface->FamilyName, self::$ExcelDigitWidthPxAt11)) {
            $pxAt11 = self::$ExcelDigitWidthPxAt11[$typeface->FamilyName];
            return true;
        }

        $pxAt11 = 0.0;
        return false;
    }

    public function MeasureDefaultRowHeightPt($normalFont)
    {
        $typeface = $this->Resolve($normalFont);
        $sizePt = $normalFont !== null ? $normalFont->Size : 11.0;
        $font = new SKFont($typeface, $sizePt);
        try {
            $metrics = $font->Metrics;
            $lineHeight = -$metrics->Ascent + $metrics->Descent + $metrics->Leading;
            if ($lineHeight <= 0.0) {
                return 14.4;
            }

            return $lineHeight + 1.0;
        } finally {
            $font->Dispose();
        }
    }

    public function Dispose()
    {
        foreach ($this->cache as $typeface) {
            if ($typeface !== null && $typeface !== SKTypeface::Default()) {
                $typeface->Dispose();
            }
        }

        $this->cache = array();

        foreach ($this->fallbacks as $typeface) {
            if ($typeface !== null) {
                $typeface->Dispose();
            }
        }

        $this->fallbacks = array();
    }
}
