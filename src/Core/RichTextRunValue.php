<?php

namespace Aspose\Cells_FOSS\Core;

final class RichTextRunValue
{
    public $StartIndex = 0;

    public $Length = 0;

    public $Font = null;

    public function __construct()
    {
        $this->Font = new FontValue();
    }

    public function Clone()
    {
        $copy = new RichTextRunValue();
        $copy->StartIndex = $this->StartIndex;
        $copy->Length = $this->Length;
        $copy->Font = $this->Font === null ? new FontValue() : $this->Font->Clone();
        return $copy;
    }
}
