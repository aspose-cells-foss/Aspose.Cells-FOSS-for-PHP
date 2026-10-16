<?php
namespace Aspose\Cells_FOSS\Core;

final class MergeRegion
{
    public $FirstRow;
    public $FirstColumn;
    public $TotalRows;
    public $TotalColumns;

    public function __construct($firstRow, $firstColumn, $totalRows, $totalColumns)
    {
        $this->FirstRow = $firstRow;
        $this->FirstColumn = $firstColumn;
        $this->TotalRows = $totalRows;
        $this->TotalColumns = $totalColumns;
    }
}
