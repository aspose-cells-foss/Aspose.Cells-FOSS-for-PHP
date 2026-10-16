<?php

namespace Aspose\Cells_FOSS\Packaging;

class MissingPartException extends PackageStructureException
{
    public function __construct($message)
    {
        parent::__construct($message);
    }
}
