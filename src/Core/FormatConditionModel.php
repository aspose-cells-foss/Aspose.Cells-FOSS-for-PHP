<?php

namespace Aspose\Cells_FOSS\Core;

use Aspose\Cells_FOSS\OperatorType;

final class FormatConditionModel
{
    public $Type = 0;
    public $Operator = OperatorType::None;
    public $Formula1;
    public $Formula2;
    public $TimePeriod;
    public $Duplicate = true;
    public $Top = true;
    public $Percent = false;
    public $Rank = 0;
    public $Above = true;
    public $StandardDeviation = 0;
    public $ColorScaleCount = 2;
    public $MinColor;
    public $MidColor;
    public $MaxColor;
    public $BarColor;
    public $NegativeBarColor;
    public $ShowBorder = false;
    public $DataBarGradient = true;
    public $DataBarMinLength = 0.10;
    public $DataBarMaxLength = 0.90;
    public $DataBarAutoMinimum = false;
    public $Direction;
    public $BarLength;
    public $IconSetType;
    public $ReverseIcons = false;
    public $ShowIconOnly = false;
    public $Priority = 0;
    public $StopIfTrue = false;
    public $Style;

    public function __construct()
    {
        $this->Style = StyleValue::Default()->Clone();
    }
}
