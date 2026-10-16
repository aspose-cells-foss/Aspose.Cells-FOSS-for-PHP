<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\SheetVisibility;
use Aspose\Cells_FOSS\Core\WorksheetModel;

class Worksheet
{
    private $_workbook;
    private $_model;
    private $_cells;
    private $_hyperlinks;
    private $_validations;
    private $_conditionalFormattings;
    private $_pageSetup;
    private $_protection;
    private $_autoFilter;
    private $_listObjects;
    private $_pictures;
    private $_shapes;
    private $_charts;
    private $_comments;

    public function __construct(Workbook $workbook, WorksheetModel $model)
    {
        $this->_workbook = $workbook;
        $this->_model = $model;
        $this->_cells = new Cells($this);
        $this->_hyperlinks = new HyperlinkCollection($model->Hyperlinks);
        $this->_validations = new ValidationCollection($model->Validations);
        $this->_conditionalFormattings = new ConditionalFormattingCollection($model->ConditionalFormattings);
        $this->_pageSetup = new PageSetup($model->PageSetup);
        $this->_protection = new WorksheetProtection($model->Protection);
        $this->_autoFilter = new AutoFilter($model->AutoFilter);
        $this->_listObjects = new ListObjectCollection($model);
        $this->_pictures = new PictureCollection($model);
        $this->_shapes = new ShapeCollection($model);
        $this->_charts = new ChartCollection($model);
        $this->_comments = new CommentCollection($model->Comments);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Model') {
            return $this->_model;
        }
        if ($propertyName === 'Workbook') {
            return $this->_workbook;
        }
        if ($propertyName === 'Name') {
            return $this->_model->Name;
        }
        if ($propertyName === 'VisibilityType') {
            switch ($this->_model->Visibility) {
                case SheetVisibility::Hidden:
                    return VisibilityType::Hidden;
                case SheetVisibility::VeryHidden:
                    return VisibilityType::VeryHidden;
                default:
                    return VisibilityType::Visible;
            }
        }
        if ($propertyName === 'TabColor') {
            if ($this->_model->TabColor !== null) {
                return Color::FromCore($this->_model->TabColor);
            }

            return Color::Empty();
        }
        if ($propertyName === 'ShowGridlines') {
            return $this->_model->View->ShowGridLines;
        }
        if ($propertyName === 'ShowRowColumnHeaders') {
            return $this->_model->View->ShowRowColumnHeaders;
        }
        if ($propertyName === 'ShowZeros') {
            return $this->_model->View->ShowZeros;
        }
        if ($propertyName === 'RightToLeft') {
            return $this->_model->View->RightToLeft;
        }
        if ($propertyName === 'Zoom') {
            return $this->_model->View->ZoomScale;
        }
        if ($propertyName === 'Cells') {
            return $this->_cells;
        }
        if ($propertyName === 'Hyperlinks') {
            return $this->_hyperlinks;
        }
        if ($propertyName === 'Validations') {
            return $this->_validations;
        }
        if ($propertyName === 'ConditionalFormattings') {
            return $this->_conditionalFormattings;
        }
        if ($propertyName === 'PageSetup') {
            return $this->_pageSetup;
        }
        if ($propertyName === 'Protection') {
            return $this->_protection;
        }
        if ($propertyName === 'AutoFilter') {
            return $this->_autoFilter;
        }
        if ($propertyName === 'ListObjects') {
            return $this->_listObjects;
        }
        if ($propertyName === 'Pictures') {
            return $this->_pictures;
        }
        if ($propertyName === 'Shapes') {
            return $this->_shapes;
        }
        if ($propertyName === 'Charts') {
            return $this->_charts;
        }
        if ($propertyName === 'Comments') {
            return $this->_comments;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Name') {
            if ($value === null || trim($value) === '') {
                throw new CellsException('Worksheet name must be non-empty.');
            }
            $this->_workbook->EnsureUniqueSheetName($value, $this->_model);
            $this->_model->Name = $value;
            return;
        }
        if ($propertyName === 'VisibilityType') {
            switch ($value) {
                case VisibilityType::Hidden:
                    $this->_model->Visibility = SheetVisibility::Hidden;
                    break;
                case VisibilityType::VeryHidden:
                    $this->_model->Visibility = SheetVisibility::VeryHidden;
                    break;
                default:
                    $this->_model->Visibility = SheetVisibility::Visible;
                    break;
            }
            return;
        }
        if ($propertyName === 'TabColor') {
            if ($value->Equals(Color::Empty())) {
                $this->_model->TabColor = null;
                return;
            }

            $this->_model->TabColor = $value->ToCore();
            return;
        }
        if ($propertyName === 'ShowGridlines') {
            $this->_model->View->ShowGridLines = $value;
            return;
        }
        if ($propertyName === 'ShowRowColumnHeaders') {
            $this->_model->View->ShowRowColumnHeaders = $value;
            return;
        }
        if ($propertyName === 'ShowZeros') {
            $this->_model->View->ShowZeros = $value;
            return;
        }
        if ($propertyName === 'RightToLeft') {
            $this->_model->View->RightToLeft = $value;
            return;
        }
        if ($propertyName === 'Zoom') {
            if ($value < 10 || $value > 400) {
                throw new CellsException('Zoom must be between 10 and 400.');
            }

            $this->_model->View->ZoomScale = $value;
            return;
        }
    }

    public function Protect()
    {
        $this->_model->Protection->IsProtected = true;
    }

    public function Unprotect()
    {
        $this->_protection->Reset();
    }
}
