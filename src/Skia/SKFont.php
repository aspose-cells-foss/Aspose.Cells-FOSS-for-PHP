<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKFont
{
    private $handle;
    private $disposed = false;
    private $typeface;
    private $size;
    private $props = array(
        'Subpixel' => false,
        'Hinting' => 0,
        'LinearMetrics' => false,
    );

    public function __construct($typeface = null, $size = 12.0)
    {
        if (!$typeface instanceof SKTypeface) {
            $typeface = SKTypeface::Default();
        }

        $this->typeface = $typeface;
        $this->size = (float)$size;
        if ($this->size <= 0.0) {
            $this->size = 12.0;
        }

        $ffi = SkiaLibrary::ffi();
        $this->handle = $ffi->sk_font_new_with_values($typeface->handle(), $this->size, 1.0, 0.0);
    }

    public function handle()
    {
        return $this->handle;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Size') {
            return $this->size;
        }

        if ($propertyName === 'Typeface') {
            return $this->typeface;
        }

        if ($propertyName === 'Metrics') {
            $ffi = SkiaLibrary::ffi();
            $metrics = $ffi->new('sk_fontmetrics_t');
            $ffi->sk_font_get_metrics($this->handle, \FFI::addr($metrics));
            return SKFontMetrics::fromNative($metrics);
        }

        if (array_key_exists($propertyName, $this->props)) {
            return $this->props[$propertyName];
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Size') {
            $this->size = (float)$value;
            if (!$this->disposed && $this->handle !== null) {
                SkiaLibrary::ffi()->sk_font_set_size($this->handle, $this->size);
            }
            return;
        }

        if ($propertyName === 'Typeface') {
            if (!$value instanceof SKTypeface) {
                $value = SKTypeface::Default();
            }

            $this->typeface = $value;
            if (!$this->disposed && $this->handle !== null) {
                SkiaLibrary::ffi()->sk_font_set_typeface($this->handle, $value->handle());
            }
            return;
        }

        if ($propertyName === 'Subpixel') {
            $this->props['Subpixel'] = (bool)$value;
            if (!$this->disposed && $this->handle !== null) {
                SkiaLibrary::ffi()->sk_font_set_subpixel($this->handle, (bool)$value);
            }
            return;
        }

        if ($propertyName === 'Hinting') {
            $this->props['Hinting'] = (int)$value;
            if (!$this->disposed && $this->handle !== null) {
                SkiaLibrary::ffi()->sk_font_set_hinting($this->handle, (int)$value);
            }
            return;
        }

        if ($propertyName === 'LinearMetrics') {
            $this->props['LinearMetrics'] = (bool)$value;
            if (!$this->disposed && $this->handle !== null) {
                SkiaLibrary::ffi()->sk_font_set_linear_metrics($this->handle, (bool)$value);
            }
            return;
        }
    }

    public function MeasureText($text, $bounds = null)
    {
        if ($this->disposed || $this->handle === null) {
            return 0.0;
        }

        $text = (string)$text;
        if ($text === '') {
            return 0.0;
        }

        $ffi = SkiaLibrary::ffi();
        $nativeBounds = $ffi->new('sk_rect_t');
        $width = $ffi->sk_font_measure_text($this->handle, $text, strlen($text), SKTextEncoding::Utf8, \FFI::addr($nativeBounds), null);
        if ($bounds instanceof SKRect) {
            $bounds->Left = $nativeBounds->left;
            $bounds->Top = $nativeBounds->top;
            $bounds->Right = $nativeBounds->right;
            $bounds->Bottom = $nativeBounds->bottom;
        }

        return $width;
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_font_delete($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
