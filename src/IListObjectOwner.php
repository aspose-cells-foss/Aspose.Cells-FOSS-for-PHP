<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ListObjectModel;

interface IListObjectOwner
{
    public function ValidateUniqueDisplayName($displayName, ListObjectModel $skipModel);
    public function ValidateNoOverlap($startRow, $startColumn, $endRow, $endColumn, ListObjectModel $skipModel);
    public function RemoveTable(ListObjectModel $model);
}