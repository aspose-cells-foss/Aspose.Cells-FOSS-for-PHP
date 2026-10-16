<?php

namespace Aspose\Cells_FOSS;

class InvalidFileFormatException extends CellsException
{
    public function __construct($message, $innerException = null)
    {
        parent::__construct($message, $innerException);
    }
}