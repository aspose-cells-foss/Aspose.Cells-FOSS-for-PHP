<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\DefinedNameModel;

final class DefinedName
{
    private $workbook;
    private $model;

    public function __construct(Workbook $workbook, DefinedNameModel $model)
    {
        $this->workbook = $workbook;
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Name') {
            return $this->model->Name;
        }
        if ($propertyName === 'Formula') {
            return $this->model->Formula;
        }
        if ($propertyName === 'LocalSheetIndex') {
            return $this->model->LocalSheetIndex;
        }
        if ($propertyName === 'Hidden') {
            return $this->model->Hidden;
        }
        if ($propertyName === 'Comment') {
            return $this->model->Comment;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Name') {
            $normalized = DefinedNameUtility::NormalizeName($value);
            $this->workbook->EnsureUniqueDefinedName($this->model, $normalized, $this->model->LocalSheetIndex);
            $this->model->Name = $normalized;
            return;
        }

        if ($propertyName === 'Formula') {
            $this->model->Formula = DefinedNameUtility::NormalizeFormula($value);
            return;
        }

        if ($propertyName === 'LocalSheetIndex') {
            $this->workbook->EnsureValidDefinedNameScope($value);
            $this->workbook->EnsureUniqueDefinedName($this->model, $this->model->Name, $value);
            $this->model->LocalSheetIndex = $value;
            return;
        }

        if ($propertyName === 'Hidden') {
            $this->model->Hidden = $value;
            return;
        }

        if ($propertyName === 'Comment') {
            $this->model->Comment = DefinedNameUtility::NormalizeComment($value);
            return;
        }
    }
}
