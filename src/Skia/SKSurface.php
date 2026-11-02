<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKSurface
{
    private $handle;
    private $canvas = null;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function CreateRaster($info)
    {
        $native = $info->toNative();
        $handle = SkiaLibrary::ffi()->sk_surface_new_raster(\FFI::addr($native), 0, null);
        return $handle === null ? null : new SKSurface($handle);
    }

    public static function Create($info)
    {
        return self::CreateRaster($info);
    }

    public function handle()
    {
        return $this->handle;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Canvas') {
            if ($this->canvas === null) {
                $this->canvas = SKCanvas::fromHandle(SkiaLibrary::ffi()->sk_surface_get_canvas($this->handle), false);
            }

            return $this->canvas;
        }

        return null;
    }

    public function Snapshot()
    {
        return SKImage::fromHandle(SkiaLibrary::ffi()->sk_surface_new_image_snapshot($this->handle));
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_surface_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
