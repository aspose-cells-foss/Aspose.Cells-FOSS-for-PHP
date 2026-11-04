<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKPaint
{
    private $handle;
    private $disposed = false;
    private $font = null;
    private $props = array(
        'Color' => null,
        'Style' => 0,
        'IsAntialias' => false,
        'StrokeWidth' => 0.0,
        'StrokeCap' => 0,
        'StrokeJoin' => 0,
        'StrokeMiter' => 4.0,
        'Shader' => null,
        'PathEffect' => null,
        'MaskFilter' => null,
        'FilterQuality' => 0,
        'TextSize' => 12.0,
        'TextAlign' => 0,
        'Typeface' => null,
    );

    public function __construct($font = null)
    {
        $this->props['Color'] = new SKColor(0, 0, 0);
        $this->handle = SkiaLibrary::ffi()->sk_paint_new();

        if ($font instanceof SKFont) {
            $this->props['Typeface'] = $font->Typeface;
            $this->props['TextSize'] = $font->Size;
            // Do not retain the external font reference: callers may dispose it.
            // fontForText() will lazily create a fresh font from the stored properties.
            $this->font = null;
            $this->syncColor();
            $this->syncStyle();
            $this->syncAntialias();
            $this->syncStrokeWidth();
        }
    }

    public function handle()
    {
        return $this->handle;
    }

    private function syncColor()
    {
        SkiaLibrary::ffi()->sk_paint_set_color($this->handle, SkiaLibrary::colorToNative($this->props['Color']));
    }

    private function syncStyle()
    {
        SkiaLibrary::ffi()->sk_paint_set_style($this->handle, $this->props['Style']);
    }

    private function syncAntialias()
    {
        SkiaLibrary::ffi()->sk_paint_set_antialias($this->handle, $this->props['IsAntialias']);
    }

    private function syncStrokeWidth()
    {
        SkiaLibrary::ffi()->sk_paint_set_stroke_width($this->handle, $this->props['StrokeWidth']);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'IsStroke') {
            return $this->props['Style'] !== SKPaintStyle::Fill;
        }

        if ($propertyName === 'FontMetrics') {
            return $this->fontForText()->Metrics;
        }

        if (array_key_exists($propertyName, $this->props)) {
            return $this->props[$propertyName];
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if (!array_key_exists($propertyName, $this->props)) {
            return;
        }

        $this->props[$propertyName] = $value;
        if ($this->disposed || $this->handle === null) {
            return;
        }

        $ffi = SkiaLibrary::ffi();

        switch ($propertyName) {
            case 'Color':
                if (!$value instanceof SKColor) {
                    $value = new SKColor(0, 0, 0);
                    $this->props['Color'] = $value;
                }
                $this->syncColor();
                break;
            case 'Style':
                $this->syncStyle();
                break;
            case 'IsAntialias':
                $this->syncAntialias();
                break;
            case 'StrokeWidth':
                $this->syncStrokeWidth();
                break;
            case 'StrokeCap':
                $ffi->sk_paint_set_stroke_cap($this->handle, (int)$value);
                break;
            case 'StrokeJoin':
                $ffi->sk_paint_set_stroke_join($this->handle, (int)$value);
                break;
            case 'StrokeMiter':
                $ffi->sk_paint_set_stroke_miter($this->handle, (float)$value);
                break;
            case 'FilterQuality':
                $ffi->sk_paint_set_filter_quality($this->handle, (int)$value);
                break;
            case 'Shader':
                $ffi->sk_paint_set_shader($this->handle, $value instanceof SKShader ? $value->handle() : null);
                break;
            case 'PathEffect':
                $ffi->sk_paint_set_path_effect($this->handle, $value instanceof SKPathEffect ? $value->handle() : null);
                break;
            case 'MaskFilter':
                $ffi->sk_paint_set_maskfilter($this->handle, $value instanceof SKMaskFilter ? $value->handle() : null);
                break;
            case 'TextSize':
            case 'Typeface':
                $this->font = null;
                break;
        }
    }

    private function fontForText()
    {
        if ($this->font !== null && !$this->font instanceof SKFont) {
            $this->font = null;
        }

        if ($this->font === null) {
            $typeface = $this->props['Typeface'];
            if (!$typeface instanceof SKTypeface) {
                $typeface = SKTypeface::Default();
            }

            $this->font = new SKFont($typeface, $this->props['TextSize']);
        }

        return $this->font;
    }

    public function MeasureText($text, $bounds = null)
    {
        return $this->fontForText()->MeasureText($text, $bounds);
    }

    public function fontHandle()
    {
        return $this->fontForText()->handle();
    }

    public function GetTextPath($text, $x, $y)
    {
        $text = (string)$text;
        if ($text === '') {
            return null;
        }

        $ffi = SkiaLibrary::ffi();
        $path = new SKPath();
        $ffi->sk_text_utils_get_path($text, strlen($text), SKTextEncoding::Utf8, (float)$x, (float)$y, $this->fontForText()->handle(), $path->handle());
        return $path;
    }

    public function ToShader()
    {
        return null;
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_paint_delete($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
