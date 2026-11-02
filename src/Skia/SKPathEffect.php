<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKPathEffect
{
    private $handle;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function CreateDash($intervals, $phase = 0.0)
    {
        $ffi = SkiaLibrary::ffi();
        $count = is_array($intervals) ? count($intervals) : 0;
        if ($count <= 0) {
            return null;
        }

        $native = $ffi->new('float[' . $count . ']');
        for ($i = 0; $i < $count; $i++) {
            $native[$i] = (float)$intervals[$i];
        }

        return new SKPathEffect($ffi->sk_path_effect_create_dash($native, $count, (float)$phase));
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
            SkiaLibrary::ffi()->sk_path_effect_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
