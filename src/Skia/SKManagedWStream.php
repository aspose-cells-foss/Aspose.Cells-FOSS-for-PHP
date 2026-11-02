<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * In-memory write stream used as the byte sink for SKDocument PDF output.
 */
final class SKManagedWStream
{
    private $handle;
    private $disposed = false;

    public function __construct()
    {
        $this->handle = SkiaLibrary::ffi()->sk_dynamicmemorywstream_new();
    }

    public function handle()
    {
        return $this->handle;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Size') {
            return (int)SkiaLibrary::ffi()->sk_wstream_bytes_written($this->castStream());
        }

        return null;
    }

    private function castStream()
    {
        return SkiaLibrary::ffi()->cast('sk_wstream_t*', $this->handle);
    }

    /**
     * Copies the buffered bytes into a PHP string.
     */
    public function ToBytes()
    {
        $ffi = SkiaLibrary::ffi();
        $size = (int)$ffi->sk_wstream_bytes_written($this->castStream());
        if ($size <= 0) {
            return '';
        }

        $buffer = $ffi->new('char[' . $size . ']');
        $ffi->sk_dynamicmemorywstream_copy_to($this->handle, $buffer);
        return \FFI::string($buffer, $size);
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_dynamicmemorywstream_destroy($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
