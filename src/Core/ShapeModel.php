<?php

namespace Aspose\Cells_FOSS\Core;

// Represents the internal model for a drawing object (shape) anchored to a worksheet.
// Row and column values are zero-based. EMU offset values are in English Metric Units.
final class ShapeModel
{
    public $Name;
    public $UpperLeftRow = 0;
    public $UpperLeftColumn = 0;
    public $UpperLeftColumnOffset = 0;
    public $UpperLeftRowOffset = 0;
    public $LowerRightRow = 0;
    public $LowerRightColumn = 0;
    public $LowerRightColumnOffset = 0;
    public $LowerRightRowOffset = 0;
    public $ExtentCx = 0;
    public $ExtentCy = 0;
    public $GeometryType;
    public $RawStyleXml;
    public $RawTxBodyXml;
    public $RawElementXml;

    public function __construct()
    {
        $this->Name = '';
        $this->GeometryType = 'rect';
    }
}