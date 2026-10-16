<?php
namespace Aspose\Cells_FOSS\Core;

final class AlignmentValue
{
    public $Horizontal;
    public $Vertical;
    public $WrapText;
    public $IndentLevel;
    public $TextRotation;
    public $ShrinkToFit;
    public $ReadingOrder;
    public $RelativeIndent;

    public function __construct()
    {
        $this->Vertical = VerticalAlignment::Bottom;
    }

    public function Clone()
    {
        $clone = new AlignmentValue();
        $clone->Horizontal = $this->Horizontal;
        $clone->Vertical = $this->Vertical;
        $clone->WrapText = $this->WrapText;
        $clone->IndentLevel = $this->IndentLevel;
        $clone->TextRotation = $this->TextRotation;
        $clone->ShrinkToFit = $this->ShrinkToFit;
        $clone->ReadingOrder = $this->ReadingOrder;
        $clone->RelativeIndent = $this->RelativeIndent;
        return $clone;
    }
}
