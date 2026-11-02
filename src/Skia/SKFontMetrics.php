<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKFontMetrics
{
    public $Flags = 0.0;
    public $Top = 0.0;
    public $Ascent = 0.0;
    public $Descent = 0.0;
    public $Bottom = 0.0;
    public $Leading = 0.0;
    public $AvgCharWidth = 0.0;
    public $MaxCharWidth = 0.0;
    public $XMin = 0.0;
    public $XMax = 0.0;
    public $XHeight = 0.0;
    public $CapHeight = 0.0;
    public $UnderlineThickness = 0.0;
    public $UnderlinePosition = 0.0;
    public $StrikeoutThickness = 0.0;
    public $StrikeoutPosition = 0.0;

    public static function fromNative($metrics)
    {
        $result = new SKFontMetrics();
        $result->Flags = $metrics->fFlags;
        $result->Top = $metrics->fTop;
        $result->Ascent = $metrics->fAscent;
        $result->Descent = $metrics->fDescent;
        $result->Bottom = $metrics->fBottom;
        $result->Leading = $metrics->fLeading;
        $result->AvgCharWidth = $metrics->fAvgCharWidth;
        $result->MaxCharWidth = $metrics->fMaxCharWidth;
        $result->XMin = $metrics->fXMin;
        $result->XMax = $metrics->fXMax;
        $result->XHeight = $metrics->fXHeight;
        $result->CapHeight = $metrics->fCapHeight;
        $result->UnderlineThickness = $metrics->fUnderlineThickness;
        $result->UnderlinePosition = $metrics->fUnderlinePosition;
        $result->StrikeoutThickness = $metrics->fStrikeoutThickness;
        $result->StrikeoutPosition = $metrics->fStrikeoutPosition;
        return $result;
    }
}
