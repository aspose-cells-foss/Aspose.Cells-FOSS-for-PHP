<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Managed wrapper over sk_canvas_t.
 */
final class SKCanvas
{
    private $handle;
    private $ownsHandle = false;
    private $disposed = false;

    public function __construct($bitmap = null)
    {
        $this->handle = $bitmap instanceof SKBitmap
            ? SkiaLibrary::ffi()->sk_canvas_new_from_bitmap($bitmap->handle())
            : null;
        $this->ownsHandle = $this->handle !== null;
    }

    public function handle()
    {
        return $this->handle;
    }

    public static function fromHandle($handle, $ownsHandle = false)
    {
        if ($handle === null) {
            return null;
        }

        $canvas = new SKCanvas();
        $canvas->handle = $handle;
        $canvas->ownsHandle = (bool)$ownsHandle;
        return $canvas;
    }

    public function Save()
    {
        return SkiaLibrary::ffi()->sk_canvas_save($this->handle);
    }

    public function Restore()
    {
        SkiaLibrary::ffi()->sk_canvas_restore($this->handle);
    }

    public function RestoreToCount($count)
    {
        SkiaLibrary::ffi()->sk_canvas_restore_to_count($this->handle, (int)$count);
    }

    public function SaveLayer($bounds = null, $paint = null)
    {
        $native = $bounds !== null ? $bounds->toNative() : null;
        return SkiaLibrary::ffi()->sk_canvas_save_layer(
            $this->handle,
            $native !== null ? \FFI::addr($native) : null,
            $paint !== null ? $paint->handle() : null
        );
    }

    public function Translate($x, $y = null)
    {
        if ($x instanceof SKPoint) {
            SkiaLibrary::ffi()->sk_canvas_translate($this->handle, $x->X, $x->Y);
            return;
        }

        SkiaLibrary::ffi()->sk_canvas_translate($this->handle, (float)$x, (float)$y);
    }

    public function Scale($scaleX, $scaleY = null)
    {
        if ($scaleY === null) {
            $scaleY = $scaleX;
        }

        SkiaLibrary::ffi()->sk_canvas_scale($this->handle, (float)$scaleX, (float)$scaleY);
    }

    public function RotateDegrees($degrees, $pivotX = null, $pivotY = null)
    {
        $ffi = SkiaLibrary::ffi();
        if ($pivotX === null && $pivotY === null) {
            $ffi->sk_canvas_rotate_degrees($this->handle, (float)$degrees);
            return;
        }

        if ($pivotX instanceof SKPoint) {
            $pivotY = $pivotX->Y;
            $pivotX = $pivotX->X;
        }

        $ffi->sk_canvas_translate($this->handle, (float)$pivotX, (float)$pivotY);
        $ffi->sk_canvas_rotate_degrees($this->handle, (float)$degrees);
        $ffi->sk_canvas_translate($this->handle, -(float)$pivotX, -(float)$pivotY);
    }

    public function Concat($matrix)
    {
        $native = $matrix->toNative();
        SkiaLibrary::ffi()->sk_canvas_concat($this->handle, \FFI::addr($native));
    }

    public function Clear($color)
    {
        SkiaLibrary::ffi()->sk_canvas_clear($this->handle, SkiaLibrary::colorToNative($color));
    }

    public function ClipRect($rect, $operation = SKClipOperation::Intersect, $antialias = false)
    {
        $native = $rect->toNative();
        SkiaLibrary::ffi()->sk_canvas_clip_rect_with_operation($this->handle, \FFI::addr($native), (int)$operation, (bool)$antialias);
    }

    public function ClipPath($path, $operation = SKClipOperation::Intersect, $antialias = false)
    {
        SkiaLibrary::ffi()->sk_canvas_clip_path_with_operation($this->handle, $path->handle(), (int)$operation, (bool)$antialias);
    }

    public function DrawRect($rectOrX, $topOrPaint = null, $width = null, $height = null, $paint = null)
    {
        $ffi = SkiaLibrary::ffi();
        if ($rectOrX instanceof SKRect) {
            $native = $rectOrX->toNative();
            $effectivePaint = $paint !== null ? $paint : $topOrPaint;
            $ffi->sk_canvas_draw_rect($this->handle, \FFI::addr($native), $effectivePaint->handle());
            return;
        }

        $native = (new SKRect($rectOrX, $topOrPaint, $rectOrX + $width, $topOrPaint + $height))->toNative();
        $ffi->sk_canvas_draw_rect($this->handle, \FFI::addr($native), $paint->handle());
    }

    public function DrawRoundRect($rect, $radiusX, $radiusY = null, $paint = null)
    {
        if ($paint === null && $radiusY instanceof SKPaint) {
            $paint = $radiusY;
            $radiusY = $radiusX;
        }

        if ($radiusY === null) {
            $radiusY = $radiusX;
        }

        $native = $rect->toNative();
        SkiaLibrary::ffi()->sk_canvas_draw_round_rect($this->handle, \FFI::addr($native), (float)$radiusX, (float)$radiusY, $paint->handle());
    }

    public function DrawLine($x1, $y1 = null, $x2 = null, $y2 = null, $paint = null)
    {
        $ffi = SkiaLibrary::ffi();
        if ($x1 instanceof SKPoint) {
            $ffi->sk_canvas_draw_line($this->handle, $x1->X, $x1->Y, $y1->X, $y1->Y, $x2->handle());
            return;
        }

        $ffi->sk_canvas_draw_line($this->handle, (float)$x1, (float)$y1, (float)$x2, (float)$y2, $paint->handle());
    }

    public function DrawCircle($centerOrX, $radiusOrY, $paintOrRadius = null, $paint = null)
    {
        $ffi = SkiaLibrary::ffi();
        if ($centerOrX instanceof SKPoint) {
            $ffi->sk_canvas_draw_circle($this->handle, $centerOrX->X, $centerOrX->Y, (float)$radiusOrY, $paintOrRadius->handle());
            return;
        }

        $ffi->sk_canvas_draw_circle($this->handle, (float)$centerOrX, (float)$radiusOrY, (float)$paintOrRadius, $paint->handle());
    }

    public function DrawOval($rect, $paint)
    {
        $native = $rect->toNative();
        SkiaLibrary::ffi()->sk_canvas_draw_oval($this->handle, \FFI::addr($native), $paint->handle());
    }

    public function DrawPath($path, $paint)
    {
        SkiaLibrary::ffi()->sk_canvas_draw_path($this->handle, $path->handle(), $paint->handle());
    }

    public function DrawImage($image, $rectOrX, $paintOrY = null, $paint = null)
    {
        $ffi = SkiaLibrary::ffi();
        if ($rectOrX instanceof SKRect) {
            $native = $rectOrX->toNative();
            $effectivePaint = $paint !== null ? $paint : $paintOrY;
            $ffi->sk_canvas_draw_image_rect($this->handle, $image->handle(), null, \FFI::addr($native), $effectivePaint->handle());
            return;
        }

        $ffi->sk_canvas_draw_image($this->handle, $image->handle(), (float)$rectOrX, (float)$paintOrY, $paint->handle());
    }

    public function DrawBitmap($bitmap, $source, $destination, $paint = null)
    {
        $ffi = SkiaLibrary::ffi();
        $image = $ffi->sk_image_new_from_bitmap($bitmap->handle());
        if ($image === null) {
            return;
        }

        try {
            $nativeSource = $source !== null ? $source->toNative() : null;
            $nativeDestination = $destination->toNative();
            $ffi->sk_canvas_draw_image_rect(
                $this->handle,
                $image,
                $nativeSource !== null ? \FFI::addr($nativeSource) : null,
                \FFI::addr($nativeDestination),
                $paint !== null ? $paint->handle() : null
            );
        } finally {
            $ffi->sk_image_unref($image);
        }
    }

    public function DrawText($text, $x, $y, $paint)
    {
        $text = (string)$text;
        if ($text === '') {
            return;
        }

        SkiaLibrary::ffi()->sk_canvas_draw_simple_text(
            $this->handle,
            $text,
            strlen($text),
            SKTextEncoding::Utf8,
            (float)$x,
            (float)$y,
            $paint->fontHandle(),
            $paint->handle()
        );
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            if ($this->ownsHandle) {
                SkiaLibrary::ffi()->sk_canvas_destroy($this->handle);
            }

            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
