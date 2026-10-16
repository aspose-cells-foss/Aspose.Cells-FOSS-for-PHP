<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\PictureModel;

// Represents a picture (image) anchored to a worksheet.
final class Picture
{
    private $_model;

    public function __construct(PictureModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Name') {
            return $this->_model->Name;
        }
        if ($propertyName === 'UpperLeftRow') {
            return $this->_model->UpperLeftRow;
        }
        if ($propertyName === 'UpperLeftColumn') {
            return $this->_model->UpperLeftColumn;
        }
        if ($propertyName === 'LowerRightRow') {
            return $this->_model->LowerRightRow;
        }
        if ($propertyName === 'LowerRightColumn') {
            return $this->_model->LowerRightColumn;
        }
        if ($propertyName === 'ImageType') {
            return self::ImageTypeFromExtension($this->_model->ImageExtension);
        }
        if ($propertyName === 'Data') {
            return $this->_model->ImageData;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Name') {
            $this->_model->Name = $value ?? '';
            return;
        }
        if ($propertyName === 'UpperLeftRow') {
            if ($value < 0) {
                throw new CellsException('UpperLeftRow must be non-negative.');
            }

            $this->_model->UpperLeftRow = $value;
            return;
        }
        if ($propertyName === 'UpperLeftColumn') {
            if ($value < 0) {
                throw new CellsException('UpperLeftColumn must be non-negative.');
            }

            $this->_model->UpperLeftColumn = $value;
            return;
        }
        if ($propertyName === 'LowerRightRow') {
            $this->_model->LowerRightRow = $value;
            return;
        }
        if ($propertyName === 'LowerRightColumn') {
            $this->_model->LowerRightColumn = $value;
            return;
        }
        if ($propertyName === 'ImageType') {
            throw new \RuntimeException('ImageType is read-only.');
        }
        if ($propertyName === 'Data') {
            if ($value === null || $value === '') {
                throw new CellsException('Picture data must be non-null and non-empty.');
            }

            $this->_model->ImageData = $value;
            $this->_model->ImageExtension = self::DetectExtension($value);
            return;
        }
    }

    public static function ImageTypeFromExtension($extension)
    {
        switch ($extension) {
            case 'jpeg':
            case 'jpg':
                return ImageType::Jpeg;
            case 'png':
                return ImageType::Png;
            case 'gif':
                return ImageType::Gif;
            case 'bmp':
                return ImageType::Bmp;
            default:
                return ImageType::Unknown;
        }
    }

    public static function ExtensionFromImageType($type)
    {
        switch ($type) {
            case ImageType::Jpeg: return 'jpeg';
            case ImageType::Png: return 'png';
            case ImageType::Gif: return 'gif';
            case ImageType::Bmp: return 'bmp';
            default: return 'bin';
        }
    }

    public static function DetectExtension($data)
    {
        if ($data === null || strlen($data) < 4) {
            return 'bin';
        }

        if (ord($data[0]) === 0xFF && ord($data[1]) === 0xD8) {
            return 'jpeg';
        }

        if (ord($data[0]) === 0x89 && ord($data[1]) === 0x50 && ord($data[2]) === 0x4E && ord($data[3]) === 0x47) {
            return 'png';
        }

        if (ord($data[0]) === 0x47 && ord($data[1]) === 0x49 && ord($data[2]) === 0x46) {
            return 'gif';
        }

        if (ord($data[0]) === 0x42 && ord($data[1]) === 0x4D) {
            return 'bmp';
        }

        return 'bin';
    }

    public static function ContentTypeFromExtension($extension)
    {
        switch ($extension) {
            case 'jpeg':
            case 'jpg':
                return 'image/jpeg';
            case 'png':
                return 'image/png';
            case 'gif':
                return 'image/gif';
            case 'bmp':
                return 'image/bmp';
            default:
                return 'application/octet-stream';
        }
    }
}