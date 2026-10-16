<?php

namespace Aspose\Cells_FOSS\Core;

use Aspose\Cells_FOSS\ChartType;

// Represents a chart embedded in a worksheet drawing layer.
final class ChartModel
{
    public $Name = '';
    public $ChartType = ChartType::Unknown;
    public $UpperLeftRow = 0;
    public $UpperLeftColumn = 0;
    public $UpperLeftRowOffset = 0;
    public $UpperLeftColumnOffset = 0;
    public $LowerRightRow = 0;
    public $LowerRightColumn = 0;
    public $LowerRightRowOffset = 0;
    public $LowerRightColumnOffset = 0;
    public $ExtentCx = 0;
    public $ExtentCy = 0;
    public $RawChartXml;
    public $RawGraphicFrameXml;
    public $OriginalRId;
    public $IsChartEx = false;
    public $CompanionFiles;

    public function __construct()
    {
        $this->CompanionFiles = array();
    }
}
