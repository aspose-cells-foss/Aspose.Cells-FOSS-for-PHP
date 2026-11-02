<?php

namespace Aspose\Cells_FOSS\Rendering;

final class Span
{
    public $Start;
    public $End;

    public function __construct($start, $end)
    {
        $this->Start = $start;
        $this->End = $end;
    }
}
