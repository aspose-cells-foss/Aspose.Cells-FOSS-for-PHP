<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKImageInfo
{
    public $Width = 0;
    public $Height = 0;
    public $ColorType = 4;
    public $AlphaType = 2;

    public function __construct($width = 0, $height = 0, $colorType = SKColorType::Rgba8888, $alphaType = SKAlphaType::Premul)
    {
        $this->Width = (int)$width;
        $this->Height = (int)$height;
        $this->ColorType = (int)$colorType;
        $this->AlphaType = (int)$alphaType;
    }

    public function toNative()
    {
        $ffi = SkiaLibrary::ffi();
        $info = $ffi->new('sk_imageinfo_t');
        $info->colorspace = null;
        $info->width = $this->Width;
        $info->height = $this->Height;
        $info->colorType = $this->ColorType;
        $info->alphaType = $this->AlphaType;
        return $info;
    }
}
