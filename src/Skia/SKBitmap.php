<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Managed wrapper over sk_bitmap_t.
 */
final class SKBitmap
{
    private $handle = null;
    private $disposed = false;

    public function __construct($width = 0, $height = 0, $colorType = null, $alphaType = null)
    {
        $ffi = SkiaLibrary::ffi();
        $this->handle = $ffi->sk_bitmap_new();

        if ($width <= 0 || $height <= 0) {
            return;
        }

        if (is_bool($colorType)) {
            $alphaType = $colorType ? SKAlphaType::Opaque : SKAlphaType::Premul;
            $colorType = SKColorType::Rgba8888;
        }

        if ($colorType === null) {
            $colorType = SKColorType::Rgba8888;
        }

        if ($alphaType === null) {
            $alphaType = SKAlphaType::Premul;
        }

        $info = new SKImageInfo($width, $height, $colorType, $alphaType);
        $native = $info->toNative();
        $ffi->sk_bitmap_try_alloc_pixels($this->handle, \FFI::addr($native), 0);
    }

    public function handle()
    {
        return $this->handle;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Width') {
            return $this->width();
        }

        if ($propertyName === 'Height') {
            return $this->height();
        }

        if ($propertyName === 'IsNull') {
            return SkiaLibrary::ffi()->sk_bitmap_is_null($this->handle);
        }

        return null;
    }

    public function width()
    {
        $ffi = SkiaLibrary::ffi();
        $info = $ffi->new('sk_imageinfo_t');
        $ffi->sk_bitmap_get_info($this->handle, \FFI::addr($info));
        return (int)$info->width;
    }

    public function height()
    {
        $ffi = SkiaLibrary::ffi();
        $info = $ffi->new('sk_imageinfo_t');
        $ffi->sk_bitmap_get_info($this->handle, \FFI::addr($info));
        return (int)$info->height;
    }

    public function GetPixel($x, $y)
    {
        return SkiaLibrary::colorFromNative(
            SkiaLibrary::ffi()->sk_bitmap_get_pixel_color($this->handle, (int)$x, (int)$y)
        );
    }

    public function SetPixel($x, $y, $color)
    {
        $addr = SkiaLibrary::ffi()->sk_bitmap_get_addr_32($this->handle, (int)$x, (int)$y);
        if ($addr === null) {
            return;
        }

        $addr[0] = SkiaLibrary::colorToNative($color);
    }

    public function ToShader($tileModeX = SKShaderTileMode::Clamp, $tileModeY = null, $matrix = null)
    {
        if ($tileModeY === null) {
            $tileModeY = $tileModeX;
        }

        $nativeMatrix = $matrix !== null ? $matrix->toNative() : null;
        $handle = SkiaLibrary::ffi()->sk_bitmap_make_shader(
            $this->handle,
            (int)$tileModeX,
            (int)$tileModeY,
            $nativeMatrix !== null ? \FFI::addr($nativeMatrix) : null
        );

        return SKShader::fromHandle($handle);
    }

    public function Copy()
    {
        $ffi = SkiaLibrary::ffi();
        $copy = new SKBitmap();
        $rect = $ffi->new('sk_irect_t');
        $rect->left = 0;
        $rect->top = 0;
        $rect->right = $this->width();
        $rect->bottom = $this->height();
        if (!$ffi->sk_bitmap_extract_subset($this->handle, $copy->handle, \FFI::addr($rect))) {
            $copy->Dispose();
            return null;
        }

        return $copy;
    }

    /**
     * Decodes an encoded image (PNG, JPEG, ...) into a bitmap. Returns null on failure.
     */
    public static function Decode($bytes)
    {
        $image = SKImage::FromEncodedData($bytes);
        if ($image === null) {
            return null;
        }

        try {
            $width = $image->Width;
            $height = $image->Height;
            if ($width <= 0 || $height <= 0) {
                return null;
            }

            $bitmap = new SKBitmap($width, $height, SKColorType::Rgba8888, SKAlphaType::Premul);
            $canvas = new SKCanvas($bitmap);
            $canvas->Clear(SKColors::Transparent());
            $paint = new SKPaint();
            $paint->IsAntialias = true;
            $paint->FilterQuality = SKFilterQuality::High;
            $canvas->DrawImage($image, new SKRect(0.0, 0.0, (float)$width, (float)$height), $paint);
            $canvas->Dispose();
            $paint->Dispose();
            return $bitmap;
        } finally {
            $image->Dispose();
        }
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_bitmap_destructor($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
