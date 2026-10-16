<?php

namespace Aspose\Cells_FOSS\Core;

use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\ValidationAlertType;

final class ValidationModel
{
    public $Areas;
    public $Type = 0;
    public $AlertStyle = ValidationAlertType::Stop;
    public $Operator = OperatorType::None;
    public $Formula1;
    public $Formula2;
    public $IgnoreBlank = false;
    public $InCellDropDown = true;
    public $InputTitle;
    public $InputMessage;
    public $ErrorTitle;
    public $ErrorMessage;
    public $ShowInput = false;
    public $ShowError = false;

    public function __construct()
    {
        $this->Areas = array();
    }
}