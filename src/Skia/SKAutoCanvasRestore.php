<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Restores the canvas save count when released (mirrors SKAutoCanvasRestore).
 */
final class SKAutoCanvasRestore
{
    private $canvas;
    private $saveCount;
    private $restored = false;

    public function __construct($canvas, $doSave = true)
    {
        $this->canvas = $canvas;
        if ($doSave) {
            $canvas->Save();
        }

        $this->saveCount = (int)SkiaLibrary::ffi()->sk_canvas_get_save_count($canvas->handle());
    }

    public static function Create($canvas, $doSave = true)
    {
        return new SKAutoCanvasRestore($canvas, $doSave);
    }

    public function Restore()
    {
        if ($this->restored || $this->canvas === null) {
            return;
        }

        $this->restored = true;
        $this->canvas->RestoreToCount($this->saveCount);
    }

    public function Dispose()
    {
        $this->Restore();
    }

    public function __destruct()
    {
        $this->Restore();
    }
}
