<?php

namespace Aspose\Cells_FOSS\Packaging;

interface IPackageWriter
{
    public function Write($stream, PackageModel $packageModel);
}
