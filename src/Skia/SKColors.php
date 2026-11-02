<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKColors
{
    private function __construct()
    {
    }

    public static function Black()
    {
        return new SKColor(0, 0, 0);
    }

    public static function White()
    {
        return new SKColor(255, 255, 255);
    }

    public static function Transparent()
    {
        return new SKColor(0, 0, 0, 0);
    }

    public static function Red()
    {
        return new SKColor(255, 0, 0);
    }

    public static function Green()
    {
        return new SKColor(0, 128, 0);
    }

    public static function Lime()
    {
        return new SKColor(0, 255, 0);
    }

    public static function Blue()
    {
        return new SKColor(0, 0, 255);
    }

    public static function Cyan()
    {
        return new SKColor(0, 255, 255);
    }

    public static function Magenta()
    {
        return new SKColor(255, 0, 255);
    }

    public static function Yellow()
    {
        return new SKColor(255, 255, 0);
    }

    public static function Gray()
    {
        return new SKColor(128, 128, 128);
    }

    public static function Silver()
    {
        return new SKColor(192, 192, 192);
    }

    public static function DarkGray()
    {
        return new SKColor(169, 169, 169);
    }

    public static function LightGray()
    {
        return new SKColor(211, 211, 211);
    }

    public static function Orange()
    {
        return new SKColor(255, 165, 0);
    }

    public static function Purple()
    {
        return new SKColor(128, 0, 128);
    }

    public static function Brown()
    {
        return new SKColor(165, 42, 42);
    }
}
