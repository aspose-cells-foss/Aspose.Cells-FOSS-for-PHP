<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKRect
{
    public $Left = 0.0;
    public $Top = 0.0;
    public $Right = 0.0;
    public $Bottom = 0.0;

    public function __construct($left = 0.0, $top = 0.0, $right = 0.0, $bottom = 0.0)
    {
        $this->Left = (float)$left;
        $this->Top = (float)$top;
        $this->Right = (float)$right;
        $this->Bottom = (float)$bottom;
    }

    public static function Create($left, $top, $right, $bottom)
    {
        return new SKRect($left, $top, $right, $bottom);
    }

    public static function CreateSize($width, $height)
    {
        return new SKRect(0.0, 0.0, $width, $height);
    }

    public static function Empty()
    {
        return new SKRect(0.0, 0.0, 0.0, 0.0);
    }

    public static function Union($first, $second)
    {
        return new SKRect(
            min($first->Left, $second->Left),
            min($first->Top, $second->Top),
            max($first->Right, $second->Right),
            max($first->Bottom, $second->Bottom)
        );
    }

    public static function Intersect($first, $second)
    {
        return new SKRect(
            max($first->Left, $second->Left),
            max($first->Top, $second->Top),
            min($first->Right, $second->Right),
            min($first->Bottom, $second->Bottom)
        );
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Width') {
            return $this->Right - $this->Left;
        }

        if ($propertyName === 'Height') {
            return $this->Bottom - $this->Top;
        }

        if ($propertyName === 'MidX') {
            return ($this->Left + $this->Right) / 2.0;
        }

        if ($propertyName === 'MidY') {
            return ($this->Top + $this->Bottom) / 2.0;
        }

        if ($propertyName === 'IsEmpty') {
            return $this->Right <= $this->Left || $this->Bottom <= $this->Top;
        }

        if ($propertyName === 'Size') {
            return new SKSize($this->Right - $this->Left, $this->Bottom - $this->Top);
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Width') {
            $this->Right = $this->Left + (float)$value;
            return;
        }

        if ($propertyName === 'Height') {
            $this->Bottom = $this->Top + (float)$value;
            return;
        }
    }

    public function Contains($x, $y = null)
    {
        if ($x instanceof SKPoint) {
            return $x->X >= $this->Left && $x->X < $this->Right && $x->Y >= $this->Top && $x->Y < $this->Bottom;
        }

        if ($x instanceof SKRect) {
            return $x->Left >= $this->Left && $x->Right <= $this->Right && $x->Top >= $this->Top && $x->Bottom <= $this->Bottom;
        }

        return $x >= $this->Left && $x < $this->Right && $y >= $this->Top && $y < $this->Bottom;
    }

    public function Inflate($dx, $dy = null)
    {
        if ($dy === null) {
            $dy = $dx;
        }

        $this->Left -= $dx;
        $this->Top -= $dy;
        $this->Right += $dx;
        $this->Bottom += $dy;
    }

    public function Offset($dx, $dy)
    {
        $this->Left += $dx;
        $this->Top += $dy;
        $this->Right += $dx;
        $this->Bottom += $dy;
    }

    public function IntersectsWith($other)
    {
        return $this->Left < $other->Right && $other->Left < $this->Right
            && $this->Top < $other->Bottom && $other->Top < $this->Bottom;
    }

    public function toNative()
    {
        $ffi = SkiaLibrary::ffi();
        $rect = $ffi->new('sk_rect_t');
        $rect->left = (float)$this->Left;
        $rect->top = (float)$this->Top;
        $rect->right = (float)$this->Right;
        $rect->bottom = (float)$this->Bottom;
        return $rect;
    }

    public static function fromNative($rect)
    {
        return new SKRect($rect->left, $rect->top, $rect->right, $rect->bottom);
    }
}
