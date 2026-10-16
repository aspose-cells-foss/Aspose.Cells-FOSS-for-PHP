<?php

namespace Aspose\Cells_FOSS\Packaging;

final class PackageLoadContext
{
    public $Workbook;
    public $Package;

    public function __construct($workbook, PackageModel $package)
    {
        $this->Workbook = $workbook;
        $this->Package = $package;
    }
}
