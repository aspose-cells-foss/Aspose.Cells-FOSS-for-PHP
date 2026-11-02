<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKPath
{
    private $handle;
    private $disposed = false;

    public function __construct()
    {
        $this->handle = SkiaLibrary::ffi()->sk_path_new();
    }

    public function handle()
    {
        return $this->handle;
    }

    public function MoveTo($x, $y = null)
    {
        if ($x instanceof SKPoint) {
            SkiaLibrary::ffi()->sk_path_move_to($this->handle, $x->X, $x->Y);
            return;
        }

        SkiaLibrary::ffi()->sk_path_move_to($this->handle, (float)$x, (float)$y);
    }

    public function LineTo($x, $y = null)
    {
        if ($x instanceof SKPoint) {
            SkiaLibrary::ffi()->sk_path_line_to($this->handle, $x->X, $x->Y);
            return;
        }

        SkiaLibrary::ffi()->sk_path_line_to($this->handle, (float)$x, (float)$y);
    }

    public function QuadTo($x1, $y1, $x2 = null, $y2 = null)
    {
        if (is_object($x1)) {
            SkiaLibrary::ffi()->sk_path_quad_to($this->handle, $x1->X, $x1->Y, $y1->X, $y1->Y);
            return;
        }

        SkiaLibrary::ffi()->sk_path_quad_to($this->handle, (float)$x1, (float)$y1, (float)$x2, (float)$y2);
    }

    public function CubicTo($x1, $y1, $x2, $y2, $x3 = null, $y3 = null)
    {
        if (is_object($x1)) {
            // SKPoint overloads are not used by the port.
            return;
        }

        SkiaLibrary::ffi()->sk_path_cubic_to($this->handle, (float)$x1, (float)$y1, (float)$x2, (float)$y2, (float)$x3, (float)$y3);
    }

    public function ConicTo($x1, $y1, $x2, $y2, $weight)
    {
        SkiaLibrary::ffi()->sk_path_conic_to($this->handle, (float)$x1, (float)$y1, (float)$x2, (float)$y2, (float)$weight);
    }

    public function Close()
    {
        SkiaLibrary::ffi()->sk_path_close($this->handle);
    }

    public function Reset()
    {
        SkiaLibrary::ffi()->sk_path_reset($this->handle);
    }

    public function Rewind()
    {
        SkiaLibrary::ffi()->sk_path_rewind($this->handle);
    }

    public function AddRect($rect)
    {
        $native = $rect->toNative();
        SkiaLibrary::ffi()->sk_path_add_rect($this->handle, \FFI::addr($native), SKPathDirection::Clockwise);
    }

    public function AddOval($rect)
    {
        $native = $rect->toNative();
        SkiaLibrary::ffi()->sk_path_add_oval($this->handle, \FFI::addr($native), SKPathDirection::Clockwise);
    }

    public function AddRoundRect($rect, $rx, $ry = null)
    {
        if ($ry === null) {
            $ry = $rx;
        }

        $native = $rect->toNative();
        SkiaLibrary::ffi()->sk_path_add_rounded_rect($this->handle, \FFI::addr($native), (float)$rx, (float)$ry, SKPathDirection::Clockwise);
    }

    public function AddCircle($cx, $cy, $radius)
    {
        SkiaLibrary::ffi()->sk_path_add_circle($this->handle, (float)$cx, (float)$cy, (float)$radius, SKPathDirection::Clockwise);
    }

    public function ArcTo($oval, $startAngle, $sweepAngle, $forceMoveTo)
    {
        $native = $oval->toNative();
        SkiaLibrary::ffi()->sk_path_arc_to_with_oval($this->handle, \FFI::addr($native), (float)$startAngle, (float)$sweepAngle, (bool)$forceMoveTo);
    }

    public function Transform($matrix)
    {
        $native = $matrix->toNative();
        SkiaLibrary::ffi()->sk_path_transform($this->handle, \FFI::addr($native));
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Bounds') {
            $ffi = SkiaLibrary::ffi();
            $bounds = $ffi->new('sk_rect_t');
            $ffi->sk_path_get_bounds($this->handle, \FFI::addr($bounds));
            return SKRect::fromNative($bounds);
        }

        if ($propertyName === 'IsEmpty') {
            return SkiaLibrary::ffi()->sk_path_count_verbs($this->handle) === 0;
        }

        return null;
    }

    public function Clone()
    {
        return SKPath::fromHandle(SkiaLibrary::ffi()->sk_path_clone($this->handle));
    }

    public static function fromHandle($handle)
    {
        $path = new SKPath();
        SkiaLibrary::ffi()->sk_path_delete($path->handle);
        $path->handle = $handle;
        return $path;
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_path_delete($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
