<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Managed wrapper over sk_image_t.
 */
final class SKImage
{
    private $handle;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function fromHandle($handle)
    {
        if ($handle === null) {
            return null;
        }

        return new SKImage($handle);
    }

    public static function FromBitmap($bitmap)
    {
        if ($bitmap === null) {
            return null;
        }

        return new SKImage(SkiaLibrary::ffi()->sk_image_new_from_bitmap($bitmap->handle()));
    }

    /**
     * Creates an image from an encoded byte string or an SKData instance.
     */
    public static function FromEncodedData($encoded)
    {
        $ffi = SkiaLibrary::ffi();
        $data = null;
        $ownsData = false;
        if ($encoded instanceof SKData) {
            $data = $encoded->handle();
        } else {
            $bytes = (string)$encoded;
            $size = strlen($bytes);
            if ($size === 0) {
                return null;
            }

            $data = $ffi->sk_data_new_with_copy($bytes, $size);
            $ownsData = true;
        }

        if ($data === null) {
            return null;
        }

        try {
            $handle = $ffi->sk_image_new_from_encoded($data, null);
        } finally {
            if ($ownsData) {
                $ffi->sk_data_unref($data);
            }
        }

        return $handle === null ? null : new SKImage($handle);
    }

    public function handle()
    {
        return $this->handle;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Width') {
            return (int)SkiaLibrary::ffi()->sk_image_get_width($this->handle);
        }

        if ($propertyName === 'Height') {
            return (int)SkiaLibrary::ffi()->sk_image_get_height($this->handle);
        }

        return null;
    }

    public function Encode($format, $quality = 100)
    {
        $handle = SkiaLibrary::ffi()->sk_image_encode_specific($this->handle, (int)$format, (int)$quality);
        return $handle === null ? null : SKData::fromHandle($handle);
    }

    public function MakeShader($tileModeX = SKShaderTileMode::Clamp, $tileModeY = null, $matrix = null, $filterQuality = SKFilterQuality::None)
    {
        if ($tileModeY === null) {
            $tileModeY = $tileModeX;
        }

        $nativeMatrix = $matrix !== null ? $matrix->toNative() : null;
        $handle = SkiaLibrary::ffi()->sk_image_make_shader(
            $this->handle,
            (int)$tileModeX,
            (int)$tileModeY,
            $nativeMatrix !== null ? \FFI::addr($nativeMatrix) : null,
            (int)$filterQuality
        );

        return SKShader::fromHandle($handle);
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_image_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
