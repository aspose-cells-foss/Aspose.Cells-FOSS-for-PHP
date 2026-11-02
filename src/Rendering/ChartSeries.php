<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class ChartSeries
{
    public $Name = '';

    public $Color;

    public $PointColors = array();

    public $LineColor;

    public $LineWidthPt = 2.25;

    public $HasVisibleLine = true;

    public $HasBarBorder = false;

    public $SmoothLine = false;

    public $Values = array();

    public $XValues = array();

    public $BubbleSizes = array();

    public $Bubble3D = false;

    public $MarkerVisible = false;

    public $MarkerSymbol = 'none';

    public $MarkerSize = 5;

    public $MarkerFillColor;

    public $MarkerStrokeColor;

    public $MarkerStrokeWidthPt = 0.75;

    public $PatternPreset;

    public $PatternForegroundColor;

    public $PatternBackgroundColor;

    public $ShowValueDataLabels = false;

    public $DataLabelPosition;

    public $DataLabelSizePt;

    public $ValueFormatCode;

    public $DataLabelColor;

    public $HasErrorBars = false;

    public $ErrorBarType;

    public $ErrorValueType;

    public $ErrorBarsHaveEndCaps = true;

    public $ErrorBarColor;

    public $ErrorBarWidthPt = 0.75;

    public function __construct()
    {
        $this->Color = SKColors::Gray();
        $this->LineColor = SKColors::Gray();
        $this->MarkerFillColor = SKColors::Gray();
        $this->MarkerStrokeColor = SKColors::Gray();
        $this->PatternForegroundColor = SKColors::Transparent();
        $this->PatternBackgroundColor = SKColors::Transparent();
        $this->DataLabelColor = new SKColor(0x40, 0x40, 0x40);
        $this->ErrorBarColor = new SKColor(0x59, 0x59, 0x59);
    }
}
