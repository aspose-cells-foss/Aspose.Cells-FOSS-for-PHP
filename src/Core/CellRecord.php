<?php
namespace Aspose\Cells_FOSS\Core;

final class CellRecord
{
    public $Value;
    public $Kind;
    public $Formula;
    public $Style;
    public $IsExplicitlyStored;

    public function __construct()
    {
        $this->Kind = CellValueKind::Blank;
        $this->Style = StyleValue::Default();
    }
}
