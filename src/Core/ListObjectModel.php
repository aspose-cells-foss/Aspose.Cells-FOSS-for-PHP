<?php

namespace Aspose\Cells_FOSS\Core;

final class ListObjectModel
{
    public $DisplayName;
    public $Name;
    public $Comment;
    public $StartRow = 0;
    public $StartColumn = 0;
    public $EndRow = 0;
    public $EndColumn = 0;
    public $ShowHeaderRow;
    public $ShowTotals = false;
    public $HasAutoFilter = false;
    public $TableStyleName;
    public $ShowFirstColumn = false;
    public $ShowLastColumn = false;
    public $ShowRowStripes;
    public $ShowColumnStripes = false;
    public $Columns;

    public function __construct()
    {
        $this->Columns = array();
        $this->DisplayName = '';
        $this->Name = '';
        $this->Comment = '';
        $this->TableStyleName = '';
        $this->ShowHeaderRow = true;
        $this->ShowRowStripes = true;
    }
}
