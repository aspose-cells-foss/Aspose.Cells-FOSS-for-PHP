<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKData;
use Aspose\Cells_FOSS\Skia\SKImage;

final class PicturePdfImageCache
{
    private $imagesByPictureVariant = array();

    private $dataByKey = array();

    private $imagesByKey = array();

    public function GetImage($picture, $widthPt, $heightPt)
    {
        if ($picture === null || $picture->ImageData === null || strlen((string)$picture->ImageData) === 0) {
            return null;
        }

        $targetWidthPx = self::ToTargetPixels($widthPt);
        $targetHeightPx = self::ToTargetPixels($heightPt);
        $variantKey = self::BuildVariantKey($picture->ImageData, $targetWidthPx, $targetHeightPx);

        if (array_key_exists($variantKey, $this->imagesByPictureVariant)) {
            return $this->imagesByPictureVariant[$variantKey];
        }

        $key = self::BuildKey($picture->ImageData);

        if (array_key_exists($key, $this->dataByKey)) {
            $data = $this->dataByKey[$key];
        } else {
            $data = SKData::CreateCopy($picture->ImageData);
            $this->dataByKey[$key] = $data;
        }

        if (array_key_exists($key, $this->imagesByKey)) {
            $cached = $this->imagesByKey[$key];
        } else {
            $cached = SKImage::FromEncodedData($data);
            $this->imagesByKey[$key] = $cached;
        }

        $this->imagesByPictureVariant[$variantKey] = $cached;
        return $cached;
    }

    private static function ToTargetPixels($points)
    {
        $pixels = $points * 96.0 / 72.0;
        if ($pixels < 1.0) {
            $pixels = 1.0;
        }

        return (int)ceil($pixels);
    }

    private static function BuildVariantKey($imageData, $targetWidthPx, $targetHeightPx)
    {
        return self::BuildKey($imageData)
            . ':'
            . $targetWidthPx
            . 'x'
            . $targetHeightPx;
    }

    private static function BuildKey($imageData)
    {
        $hash = 2166136261;
        $length = strlen((string)$imageData);
        for ($i = 0; $i < $length; $i++) {
            $hash ^= ord($imageData[$i]);
            $hash = ($hash * 16777619) & 0xFFFFFFFF;
        }

        return $length . ':' . strtoupper(str_pad(dechex($hash), 8, '0', STR_PAD_LEFT));
    }

    public function Dispose()
    {
        foreach ($this->imagesByKey as $image) {
            if ($image !== null) {
                $image->Dispose();
            }
        }

        foreach ($this->dataByKey as $data) {
            if ($data !== null) {
                $data->Dispose();
            }
        }

        $this->imagesByPictureVariant = array();
        $this->imagesByKey = array();
        $this->dataByKey = array();
    }
}
