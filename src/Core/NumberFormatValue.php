<?php
namespace Aspose\Cells_FOSS\Core;

final class NumberFormatValue
{
    public $Number;
    public $Custom;

    public function Clone()
    {
        $clone = new NumberFormatValue();
        $clone->Number = $this->Number;
        $clone->Custom = $this->Custom;
        return $clone;
    }
}
