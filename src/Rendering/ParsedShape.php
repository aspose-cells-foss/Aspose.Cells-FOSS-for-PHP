<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

// A drawing shape reduced to what the renderer needs, parsed from the preserved shape XML
// (xdr:sp / xdr:cxnSp): geometry, fill, line (width + arrowheads), and text.
final class ParsedShape
{
    public $Geometry = 'rect';
    public $FlipH = false;
    public $FlipV = false;
    public $RotationDeg = 0.0;
    public $Adjusts;   // preset geometry adjust values, as fractions (0..1)
    public $CustomGeometry;

    public $HasFill = false;
    public $FillColor;
    public $FillPattern;
    public $PatternForeground;
    public $Reflection;

    public $HasLine = true;
    public $LineColor;
    public $LineWidthPt = 0.75;

    public $HeadEnd;   // arrowhead at the start (null/none)
    public $TailEnd;   // arrowhead at the end

    public $Text;
    public $TextColor;
    public $TextSizePt = 11.0;
    public $TextBold = false;
    public $TextVAnchor = 'ctr';  // t / ctr / b
    public $TextHAlign = 'ctr';   // l / ctr / r
    public $TextVerticalType = 'horz';
    public $TextWrap = false;
    public $TextInsetTopPt = 0.0;
    public $TextInsetLeftPt = 0.0;

    public function __construct()
    {
        $this->FillColor = SKColors::White();
        $this->PatternForeground = new SKColor(0, 0, 0, 0);
        $this->LineColor = SKColors::Black();
        $this->TextColor = SKColors::Black();
    }
}
