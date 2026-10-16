<?php
namespace Aspose\Cells_FOSS\Core;

final class BordersValue
{
    public $Left;
    public $Right;
    public $Top;
    public $Bottom;
    public $Diagonal;
    public $DiagonalUp;
    public $DiagonalDown;

    public function __construct()
    {
        $this->Left = new BorderSideValue();
        $this->Right = new BorderSideValue();
        $this->Top = new BorderSideValue();
        $this->Bottom = new BorderSideValue();
        $this->Diagonal = new BorderSideValue();
    }

    public function Clone()
    {
        $clone = new BordersValue();
        $clone->Left = $this->Left->Clone();
        $clone->Right = $this->Right->Clone();
        $clone->Top = $this->Top->Clone();
        $clone->Bottom = $this->Bottom->Clone();
        $clone->Diagonal = $this->Diagonal->Clone();
        $clone->DiagonalUp = $this->DiagonalUp;
        $clone->DiagonalDown = $this->DiagonalDown;
        return $clone;
    }
}
