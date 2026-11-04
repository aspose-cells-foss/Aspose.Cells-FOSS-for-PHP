<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKPathIterator
{
    private $handle;
    private $disposed = false;

    public function __construct($pathHandle, $forceClose = false)
    {
        $this->handle = SkiaLibrary::ffi()->sk_path_create_iter($pathHandle, (bool)$forceClose);
    }

    /**
     * Advances to the next verb and fills $points (an array of up to 4 SKPoint).
     * Returns the verb (SKPathVerb constant).
     */
    public function Next(&$points)
    {
        if ($this->disposed || $this->handle === null) {
            $points = array();
            return SKPathVerb::Done;
        }

        $ffi = SkiaLibrary::ffi();
        $nativePoints = $ffi->new('sk_point_t[4]');
        $verb = $ffi->sk_path_iter_next($this->handle, $nativePoints);

        $points = array(
            new SKPoint((float)$nativePoints[0]->x, (float)$nativePoints[0]->y),
            new SKPoint((float)$nativePoints[1]->x, (float)$nativePoints[1]->y),
            new SKPoint((float)$nativePoints[2]->x, (float)$nativePoints[2]->y),
            new SKPoint((float)$nativePoints[3]->x, (float)$nativePoints[3]->y),
        );

        return (int)$verb;
    }

    public function ConicWeight()
    {
        if ($this->disposed || $this->handle === null) {
            return 0.0;
        }

        return (float)SkiaLibrary::ffi()->sk_path_iter_conic_weight($this->handle);
    }

    public function IsCloseLine()
    {
        if ($this->disposed || $this->handle === null) {
            return false;
        }

        return (bool)SkiaLibrary::ffi()->sk_path_iter_is_close_line($this->handle);
    }

    public function IsClosedContour()
    {
        if ($this->disposed || $this->handle === null) {
            return false;
        }

        return (bool)SkiaLibrary::ffi()->sk_path_iter_is_closed_contour($this->handle);
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_path_iter_destroy($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
