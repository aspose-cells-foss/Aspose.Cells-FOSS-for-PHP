<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKMaskFilter
{
    private $handle;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function CreateBlur($blurStyle, $sigma)
    {
        return new SKMaskFilter(SkiaLibrary::ffi()->sk_maskfilter_new_blur((int)$blurStyle, (float)$sigma));
    }

    public function handle()
    {
        return $this->handle;
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_maskfilter_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
