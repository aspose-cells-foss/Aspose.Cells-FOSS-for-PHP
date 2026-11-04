<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKTypeface;

final class PdfTrueTypePrototypeUsage
{
    private $typeface;
    private $widths1000 = array();
    private $codeToGlyphId = array();
    private $glyphIdToUnicode = array();
    private $codes = array();
    private $font;
    private $paint;
    private $fontBytes;
    private $subsetProgram;
    private $fontBounds;
    private $hasBounds = false;

    public $ResourceName = '';
    public $FontObjectId = 0;
    public $FontDescriptorObjectId = 0;
    public $FontFileObjectId = 0;
    public $ToUnicodeObjectId = 0;
    public $SubsetFontName = '';

    public function __construct($typeface)
    {
        if ($typeface === null) {
            throw new \InvalidArgumentException('typeface');
        }

        $this->typeface = $typeface;
        $this->font = new SKFont($typeface, 1000.0);
        $this->font->Subpixel = true;
        $this->paint = new SKPaint($this->font);
        $this->paint->Style = SKPaintStyle::Fill;
        $this->paint->IsAntialias = true;

        $this->fontBytes = TrueTypeFontBytesBuilder::Build($typeface);
        if ($this->fontBytes === '' || strlen($this->fontBytes) < 12) {
            throw new \RuntimeException('Could not read font bytes for ' . $typeface->FamilyName . '.');
        }

        $this->subsetProgram = new TrueTypeSubsetFontProgram($this->fontBytes);
    }

    public function Typeface()
    {
        return $this->typeface;
    }

    public function FontBytes()
    {
        return $this->fontBytes;
    }

    public function SubsetProgram()
    {
        return $this->subsetProgram;
    }

    public function BaseFontName()
    {
        $family = $this->typeface->FamilyName;
        if ($family === null || $family === '') {
            $family = 'SkiaFont';
        }

        $builder = '';
        $len = mb_strlen($family, 'UTF-8');
        for ($i = 0; $i < $len; $i++) {
            $ch = mb_substr($family, $i, 1, 'UTF-8');
            if (ctype_alnum($ch)) {
                $builder .= $ch;
            }
        }

        if ($builder === '') {
            $builder = 'SkiaFont';
        }

        return $builder;
    }

    public function FirstChar()
    {
        $first = 255;
        foreach ($this->codes as $code) {
            if ($code < $first) {
                $first = $code;
            }
        }

        if (count($this->codes) === 0) {
            return 32;
        }

        return $first;
    }

    public function LastChar()
    {
        $last = 32;
        foreach ($this->codes as $code) {
            if ($code > $last) {
                $last = $code;
            }
        }

        return $last;
    }

    public function FontBounds()
    {
        return $this->fontBounds;
    }

    public function HasBounds()
    {
        return $this->hasBounds;
    }

    public function AddText($text)
    {
        if ($text === null || $text === '') {
            return;
        }

        $len = strlen($text);
        for ($i = 0; $i < $len; $i++) {
            $ch = ord($text[$i]);
            if ($ch > 255) {
                continue;
            }

            $code = $ch;
            if (!isset($this->codes[$code])) {
                $this->codes[$code] = true;
                $token = $text[$i];
                $glyphId = $this->ResolveGlyphId($ch);
                $this->codeToGlyphId[$code] = $glyphId;
                if (!isset($this->glyphIdToUnicode[$glyphId])) {
                    $this->glyphIdToUnicode[$glyphId] = $ch;
                }

                $this->widths1000[$code] = $this->paint->MeasureText($token);
                $path = $this->paint->GetTextPath($token, 0.0, 0.0);
                if ($path !== null && !$path->IsEmpty) {
                    $bounds = $path->Bounds;
                    if (!$this->hasBounds) {
                        $this->fontBounds = $bounds;
                        $this->hasBounds = true;
                    } else {
                        $this->fontBounds = \Aspose\Cells_FOSS\Skia\SKRect::Union($this->fontBounds, $bounds);
                    }
                }

                if ($path !== null) {
                    $path->Dispose();
                }
            }
        }
    }

    public function EncodeText($text)
    {
        if ($text === null || $text === '') {
            return '';
        }

        $bytes = '';
        $len = strlen($text);
        for ($i = 0; $i < $len; $i++) {
            $ch = ord($text[$i]);
            if ($ch > 255) {
                continue;
            }

            $bytes .= chr($ch);
        }

        return $bytes;
    }

    public function WidthForCode($code)
    {
        if (isset($this->widths1000[$code])) {
            return $this->widths1000[$code];
        }

        return 0.0;
    }

    public function OrderedCodes()
    {
        $ordered = array_keys($this->codes);
        sort($ordered, SORT_NUMERIC);
        return $ordered;
    }

    public function OrderedGlyphIds()
    {
        $glyphIds = array_values($this->codeToGlyphId);
        sort($glyphIds, SORT_NUMERIC);
        return $glyphIds;
    }

    public function CodeToGlyphIdMap()
    {
        return $this->codeToGlyphId;
    }

    public function GlyphIdToUnicodeMap()
    {
        return $this->glyphIdToUnicode;
    }

    public function BuildSubsetFontBytes()
    {
        return $this->subsetProgram->BuildSubset(
            $this->OrderedGlyphIds(),
            $this->codeToGlyphId,
            $this->glyphIdToUnicode,
            $this->SubsetFontName
        );
    }

    private function ResolveGlyphId($ch)
    {
        return $this->font->GetGlyph($ch);
    }

    public function Dispose()
    {
        if ($this->paint !== null) {
            $this->paint->Dispose();
            $this->paint = null;
        }

        if ($this->font !== null) {
            $this->font->Dispose();
            $this->font = null;
        }
    }
}
