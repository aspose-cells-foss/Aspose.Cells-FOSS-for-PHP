<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKTypeface;

final class PdfType3FontUsage
{
    private $typeface;
    private $usePositiveYFontMatrix;
    private $tokenWidths = array();
    private $subsets = array();
    private $paint;

    public function __construct($typeface, $usePositiveYFontMatrix)
    {
        $this->typeface = $typeface;
        $this->usePositiveYFontMatrix = (bool)$usePositiveYFontMatrix;

        $font = new SKFont($typeface, 1000.0);
        $font->Subpixel = true;
        $this->paint = new SKPaint($font);
        $this->paint->Style = SKPaintStyle::Fill;
        $this->paint->IsAntialias = true;
        $font->Dispose();
    }

    public function Typeface()
    {
        return $this->typeface;
    }

    public function UsePositiveYFontMatrix()
    {
        return $this->usePositiveYFontMatrix;
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

    /**
     * @return PdfType3FontSubset[]
     */
    public function &Subsets()
    {
        return $this->subsets;
    }

    public function AddText($text)
    {
        if ($text === null || $text === '') {
            return;
        }

        $len = mb_strlen($text, 'UTF-8');
        for ($i = 0; $i < $len; $i++) {
            $token = mb_substr($text, $i, 1, 'UTF-8');
            if (!array_key_exists($token, $this->tokenWidths)) {
                $this->tokenWidths[$token] = $this->paint->MeasureText($token);
            }
        }
    }

    public function FinalizeSubsets(&$resourceIndex)
    {
        if (count($this->subsets) > 0) {
            return;
        }

        $orderedTokens = array_keys($this->tokenWidths);
        sort($orderedTokens, SORT_STRING);

        $current = null;
        for ($i = 0; $i < count($orderedTokens); $i++) {
            $token = $orderedTokens[$i];
            if ($current === null || count($current->Tokens) >= 255) {
                $current = new PdfType3FontSubset();
                $current->ResourceName = 'T3F' . $resourceIndex;
                $resourceIndex++;
                $this->subsets[] = $current;
            }

            $code = (int)(count($current->Tokens) + 1);
            $current->Tokens[] = $token;
            $current->TokenCodes[$token] = $code;
        }
    }

    public function TryResolveToken($token, &$subset, &$code, &$width1000)
    {
        $width1000 = 0.0;
        $subset = null;
        $code = 0;

        if (!array_key_exists($token, $this->tokenWidths)) {
            return false;
        }

        $width1000 = $this->tokenWidths[$token];
        for ($i = 0; $i < count($this->subsets); $i++) {
            if (array_key_exists($token, $this->subsets[$i]->TokenCodes)) {
                $subset = $this->subsets[$i];
                $code = $this->subsets[$i]->TokenCodes[$token];
                return true;
            }
        }

        return false;
    }

    public function GetTokenWidth($token)
    {
        if (array_key_exists($token, $this->tokenWidths)) {
            return $this->tokenWidths[$token];
        }

        return 0.0;
    }

    public function BuildTokenPath($token)
    {
        return $this->paint->GetTextPath($token, 0.0, 0.0);
    }

    public function Dispose()
    {
        if ($this->paint !== null) {
            $this->paint->Dispose();
            $this->paint = null;
        }
    }
}
