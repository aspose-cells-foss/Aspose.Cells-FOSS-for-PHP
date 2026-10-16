<?php

namespace Aspose\Cells_FOSS;

// Represents borders.
class Borders
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
        $this->Left = new Border();
        $this->Right = new Border();
        $this->Top = new Border();
        $this->Bottom = new Border();
        $this->Diagonal = new Border();
    }

    public function Clone()
    {
        $clone = new Borders();
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
