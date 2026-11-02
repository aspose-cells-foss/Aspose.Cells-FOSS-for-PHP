<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Owns a native sk_data_t reference.
 */
final class SKData
{
    private $handle;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function fromHandle($handle)
    {
        return new SKData($handle);
    }

    public static function CreateCopy($bytes)
    {
        $ffi = SkiaLibrary::ffi();
        return new SKData($ffi->sk_data_new_with_copy((string)$bytes, strlen((string)$bytes)));
    }

    public function handle()
    {
        return $this->handle;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Size') {
            return (int)SkiaLibrary::ffi()->sk_data_get_size($this->handle);
        }

        return null;
    }

    public function ToBytes()
    {
        $size = (int)SkiaLibrary::ffi()->sk_data_get_size($this->handle);
        if ($size <= 0) {
            return '';
        }

        $ffi = SkiaLibrary::ffi();
        $bytes = $ffi->sk_data_get_bytes($this->handle);
        if ($bytes === null) {
            return '';
        }

        return \FFI::string($bytes, $size);
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_data_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
