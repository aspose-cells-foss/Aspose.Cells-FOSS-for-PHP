<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\HyperlinkModel;

class Hyperlink
{
    private $owner;
    private $model;

    public function __construct(&$owner, HyperlinkModel $model)
    {
        $this->owner = &$owner;
        $this->model = $model;
    }

    public function BelongsTo($owner)
    {
        return $this->owner === $owner;
    }

    public function GetModel()
    {
        return $this->model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Area') {
            $first = (new CellAddress($this->model->FirstRow, $this->model->FirstColumn))->ToString();
            if ($this->model->TotalRows === 1 && $this->model->TotalColumns === 1) {
                return $first;
            }

            $last = (new CellAddress($this->model->FirstRow + $this->model->TotalRows - 1, $this->model->FirstColumn + $this->model->TotalColumns - 1))->ToString();
            return $first . ':' . $last;
        }

        if ($propertyName === 'Address') {
            $address = $this->model->Address;
            if (!($address === null || $address === '')) {
                return $address;
            }

            return $this->model->SubAddress === null ? '' : $this->model->SubAddress;
        }

        if ($propertyName === 'LinkType') {
            if (!($this->model->SubAddress === null || $this->model->SubAddress === '')) {
                return TargetModeType::CellReference;
            }

            $address = $this->model->Address;
            if ($address === null || $address === '') {
                return TargetModeType::External;
            }

            if (str_starts_with(strtolower($address), 'mailto:')) {
                return TargetModeType::Email;
            }

            if (str_starts_with($address, '\\') || strpos($address, ':\\') > 0) {
                return TargetModeType::FilePath;
            }

            return TargetModeType::External;
        }

        if ($propertyName === 'ScreenTip') {
            return $this->model->ScreenTip === null ? '' : $this->model->ScreenTip;
        }

        if ($propertyName === 'TextToDisplay') {
            if (!($this->model->TextToDisplay === null || $this->model->TextToDisplay === '')) {
                return $this->model->TextToDisplay;
            }

            return $this->__get('Address');
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Address') {
            $this->AssignAddress($value);
            return;
        }

        if ($propertyName === 'ScreenTip') {
            $this->model->ScreenTip = self::NormalizeText($value);
            return;
        }

        if ($propertyName === 'TextToDisplay') {
            $this->model->TextToDisplay = self::NormalizeText($value);
            return;
        }
    }

    public function Delete()
    {
        for ($index = 0; $index < count($this->owner); $index++) {
            if ($this->owner[$index] === $this->model) {
                array_splice($this->owner, $index, 1);
                return;
            }
        }
    }

    private function AssignAddress($value)
    {
        if ($value === null || trim($value) === '') {
            $this->model->Address = null;
            $this->model->SubAddress = null;
            return;
        }

        $normalized = trim($value);
        if (str_starts_with($normalized, '#')) {
            $this->model->Address = null;
            $this->model->SubAddress = substr($normalized, 1);
            return;
        }

        if (strpos($normalized, '!') !== false) {
            $this->model->Address = null;
            $this->model->SubAddress = $normalized;
            return;
        }

        $this->model->Address = $normalized;
        $this->model->SubAddress = null;
    }

    private static function NormalizeText($value)
    {
        if ($value === null || $value === '') {
            return null;
        }

        return $value;
    }
}
