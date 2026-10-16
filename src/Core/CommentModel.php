<?php

namespace Aspose\Cells_FOSS\Core;

// Represents the internal model for a worksheet comment (legacy note).
// Row and column values are zero-based.
final class CommentModel
{
    public $Row = 0;
    public $Column = 0;
    public $Author;
    public $Note;
    public $IsVisible = false;
    public $Width;
    public $Height;
    public $RawVmlShapeXml;

    public function __construct()
    {
        $this->Author = '';
        $this->Note = '';
        $this->Width = 129;
        $this->Height = 75;
    }
}
