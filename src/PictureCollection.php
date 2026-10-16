<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\PictureModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;

// Represents collection of pictures anchored to a worksheet.
final class PictureCollection implements \ArrayAccess, \Countable
{
    private $_worksheetModel;
    private $_models;

    public function __construct(WorksheetModel $worksheetModel)
    {
        $this->_worksheetModel = $worksheetModel;
        $this->_models =& $worksheetModel->Pictures;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_models);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Count') {
            throw new \RuntimeException('Count is read-only.');
        }
    }

    public function count(): int
    {
        return count($this->_models);
    }

    public function offsetExists($offset): bool
    {
        $index = (int) $offset;
        return $index >= 0 && $index < count($this->_models);
    }

    public function offsetGet($offset): mixed
    {
        $index = (int) $offset;
        if ($index < 0 || $index >= count($this->_models)) {
            throw new CellsException('Picture index ' . $index . ' is out of range.');
        }

        return new Picture($this->_models[$index]);
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The picture collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The picture collection is read-only.');
    }

    public function Add($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn, $imageDataOrStreamOrFileName)
    {
        if (is_resource($imageDataOrStreamOrFileName)) {
            $stream = $imageDataOrStreamOrFileName;
            $data = stream_get_contents($stream);
            return $this->Add($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn, $data);
        }

        if (is_string($imageDataOrStreamOrFileName) && is_file($imageDataOrStreamOrFileName)) {
            $fileName = $imageDataOrStreamOrFileName;
            if ($fileName === null || $fileName === '') {
                throw new CellsException('File name must be non-empty.');
            }

            $data = file_get_contents($fileName);
            $model = $this->CreateModel($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn, $data);

            $ext = '.' . pathinfo($fileName, PATHINFO_EXTENSION);
            if (!($ext === null || $ext === '') && str_starts_with($ext, '.')) {
                $ext = strtolower(substr($ext, 1));
                if ($ext === 'jpg') {
                    $ext = 'jpeg';
                }

                $model->ImageExtension = $ext;
            }

            $this->_models[] = $model;
            return count($this->_models) - 1;
        }

        $imageData = $imageDataOrStreamOrFileName;
        self::ValidateAnchor($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn);
        if ($imageData === null || $imageData === '') {
            throw new CellsException('Image data must be non-null and non-empty.');
        }

        $model = $this->CreateModel($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn, $imageData);
        $this->_models[] = $model;
        return count($this->_models) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->_models)) {
            throw new CellsException('Picture index ' . $index . ' is out of range.');
        }

        array_splice($this->_models, $index, 1);
    }

    private function CreateModel($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn, $data)
    {
        self::ValidateAnchor($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn);
        $pictureNumber = count($this->_models) + 1;
        $model = new PictureModel();
        $model->Name = 'Picture ' . $pictureNumber;
        $model->UpperLeftRow = $upperLeftRow;
        $model->UpperLeftColumn = $upperLeftColumn;
        $model->LowerRightRow = $lowerRightRow;
        $model->LowerRightColumn = $lowerRightColumn;
        $model->ImageData = $data;
        $model->ImageExtension = Picture::DetectExtension($data);
        return $model;
    }

    private static function ValidateAnchor($upperLeftRow, $upperLeftColumn, $lowerRightRow, $lowerRightColumn)
    {
        if ($upperLeftRow < 0) {
            throw new CellsException('UpperLeftRow must be non-negative.');
        }

        if ($upperLeftColumn < 0) {
            throw new CellsException('UpperLeftColumn must be non-negative.');
        }

        if ($lowerRightRow < $upperLeftRow) {
            throw new CellsException('LowerRightRow must be greater than or equal to UpperLeftRow.');
        }

        if ($lowerRightColumn < $upperLeftColumn) {
            throw new CellsException('LowerRightColumn must be greater than or equal to UpperLeftColumn.');
        }
    }
}