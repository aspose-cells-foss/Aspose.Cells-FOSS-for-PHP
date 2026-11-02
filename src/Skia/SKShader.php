<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKShader
{
    private $handle;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function CreateColor($color)
    {
        return new SKShader(SkiaLibrary::ffi()->sk_shader_new_color(SkiaLibrary::colorToNative($color)));
    }

    public static function CreateLinearGradient($start, $end, $colors, $colorPos = null, $tileMode = SKShaderTileMode::Clamp, $localMatrix = null)
    {
        $ffi = SkiaLibrary::ffi();
        $points = $ffi->new('sk_point_t[2]');
        $points[0]->x = $start->X;
        $points[0]->y = $start->Y;
        $points[1]->x = $end->X;
        $points[1]->y = $end->Y;

        $count = count($colors);
        $nativeColors = $ffi->new('uint32_t[' . ($count <= 0 ? 1 : $count) . ']');
        for ($i = 0; $i < $count; $i++) {
            $nativeColors[$i] = SkiaLibrary::colorToNative($colors[$i]);
        }

        $nativePositions = null;
        if (is_array($colorPos) && count($colorPos) === $count) {
            $nativePositions = $ffi->new('float[' . $count . ']');
            for ($i = 0; $i < $count; $i++) {
                $nativePositions[$i] = (float)$colorPos[$i];
            }
        }

        $nativeMatrix = $localMatrix instanceof SKMatrix ? $localMatrix->toNative() : null;

        return new SKShader($ffi->sk_shader_new_linear_gradient(
            $points,
            $nativeColors,
            $nativePositions === null ? null : $nativePositions,
            $count,
            (int)$tileMode,
            $nativeMatrix === null ? null : $nativeMatrix));
    }

    public static function CreateRadialGradient($center, $radius, $colors, $colorPos = null, $tileMode = SKShaderTileMode::Clamp, $localMatrix = null)
    {
        $ffi = SkiaLibrary::ffi();
        $nativeCenter = $ffi->new('sk_point_t');
        $nativeCenter->x = $center->X;
        $nativeCenter->y = $center->Y;

        $count = count($colors);
        $nativeColors = $ffi->new('uint32_t[' . ($count <= 0 ? 1 : $count) . ']');
        for ($i = 0; $i < $count; $i++) {
            $nativeColors[$i] = SkiaLibrary::colorToNative($colors[$i]);
        }

        $nativePositions = null;
        if (is_array($colorPos) && count($colorPos) === $count) {
            $nativePositions = $ffi->new('float[' . $count . ']');
            for ($i = 0; $i < $count; $i++) {
                $nativePositions[$i] = (float)$colorPos[$i];
            }
        }

        $nativeMatrix = $localMatrix instanceof SKMatrix ? $localMatrix->toNative() : null;

        return new SKShader($ffi->sk_shader_new_radial_gradient(
            $nativeCenter,
            (float)$radius,
            $nativeColors,
            $nativePositions === null ? null : $nativePositions,
            $count,
            (int)$tileMode,
            $nativeMatrix === null ? null : $nativeMatrix));
    }

    public function handle()
    {
        return $this->handle;
    }

    public static function fromHandle($handle)
    {
        if ($handle === null) {
            return null;
        }

        return new SKShader($handle);
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_shader_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
