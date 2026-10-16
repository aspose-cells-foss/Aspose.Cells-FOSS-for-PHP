<?php
namespace Aspose\Cells_FOSS\Core;

use Aspose\Cells_FOSS\FontUnderlineType;

final class FontValue
{
    public $Name;
    public $Size;
    public $Bold;
    public $Italic;
    public $Underline;
    public $StrikeThrough;
    public $Color;

    public function __construct()
    {
        $this->Name = "Calibri";
        $this->Size = 11.0;
        $this->Bold = false;
        $this->Italic = false;
        $this->Underline = FontUnderlineType::None;
        $this->StrikeThrough = false;
        $this->Color = new ColorValue();
    }

    public function Clone()
    {
        $clone = new FontValue();
        $clone->Name = $this->Name;
        $clone->Size = $this->Size;
        $clone->Bold = $this->Bold;
        $clone->Italic = $this->Italic;
        $clone->Underline = $this->Underline;
        $clone->StrikeThrough = $this->StrikeThrough;
        $clone->Color = ($this->Color === null) ? new ColorValue() : clone $this->Color;
        return $clone;
    }
}
