<?php

namespace Aspose\Cells_FOSS;

class CellsException extends \Exception
{
    public function __construct($message, $innerException = null)
    {
        parent::__construct($message, 0, $innerException);
    }
}