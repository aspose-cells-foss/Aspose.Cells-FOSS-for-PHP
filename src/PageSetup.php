<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\PageOrientation;
use Aspose\Cells_FOSS\Core\PageSetupModel;

// Represents worksheet print and page-layout settings.
final class PageSetup
{
    private const CentimetersPerInch = 2.54;
    private $_model;

    public function __construct(PageSetupModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'PaperSize') {
            return $this->_model->PaperSize;
        }
        if ($propertyName === 'Orientation') {
            switch ($this->_model->Orientation) {
                case PageOrientation::Portrait:
                    return PageOrientationType::Portrait;
                case PageOrientation::Landscape:
                    return PageOrientationType::Landscape;
                default:
                    return PageOrientationType::Default;
            }
        }
        if ($propertyName === 'FirstPageNumber') {
            return $this->_model->FirstPageNumber;
        }
        if ($propertyName === 'Scale') {
            return $this->_model->Scale;
        }
        if ($propertyName === 'FitToPagesWide') {
            return $this->_model->FitToWidth;
        }
        if ($propertyName === 'FitToPagesTall') {
            return $this->_model->FitToHeight;
        }
        if ($propertyName === 'PrintArea') {
            return $this->_model->PrintArea ?? '';
        }
        if ($propertyName === 'PrintTitleRows') {
            return $this->_model->PrintTitleRows ?? '';
        }
        if ($propertyName === 'PrintTitleColumns') {
            return $this->_model->PrintTitleColumns ?? '';
        }
        if ($propertyName === 'LeftMargin') {
            return self::ToCentimeters($this->_model->Margins->Left);
        }
        if ($propertyName === 'RightMargin') {
            return self::ToCentimeters($this->_model->Margins->Right);
        }
        if ($propertyName === 'TopMargin') {
            return self::ToCentimeters($this->_model->Margins->Top);
        }
        if ($propertyName === 'BottomMargin') {
            return self::ToCentimeters($this->_model->Margins->Bottom);
        }
        if ($propertyName === 'HeaderMargin') {
            return self::ToCentimeters($this->_model->Margins->Header);
        }
        if ($propertyName === 'FooterMargin') {
            return self::ToCentimeters($this->_model->Margins->Footer);
        }
        if ($propertyName === 'LeftMarginInch') {
            return $this->_model->Margins->Left;
        }
        if ($propertyName === 'RightMarginInch') {
            return $this->_model->Margins->Right;
        }
        if ($propertyName === 'TopMarginInch') {
            return $this->_model->Margins->Top;
        }
        if ($propertyName === 'BottomMarginInch') {
            return $this->_model->Margins->Bottom;
        }
        if ($propertyName === 'HeaderMarginInch') {
            return $this->_model->Margins->Header;
        }
        if ($propertyName === 'FooterMarginInch') {
            return $this->_model->Margins->Footer;
        }
        if ($propertyName === 'LeftHeader') {
            return $this->_model->HeaderFooter->LeftHeader ?? '';
        }
        if ($propertyName === 'CenterHeader') {
            return $this->_model->HeaderFooter->CenterHeader ?? '';
        }
        if ($propertyName === 'RightHeader') {
            return $this->_model->HeaderFooter->RightHeader ?? '';
        }
        if ($propertyName === 'LeftFooter') {
            return $this->_model->HeaderFooter->LeftFooter ?? '';
        }
        if ($propertyName === 'CenterFooter') {
            return $this->_model->HeaderFooter->CenterFooter ?? '';
        }
        if ($propertyName === 'RightFooter') {
            return $this->_model->HeaderFooter->RightFooter ?? '';
        }
        if ($propertyName === 'PrintGridlines') {
            return $this->_model->PrintOptions->GridLines;
        }
        if ($propertyName === 'PrintHeadings') {
            return $this->_model->PrintOptions->Headings;
        }
        if ($propertyName === 'CenterHorizontally') {
            return $this->_model->PrintOptions->HorizontalCentered;
        }
        if ($propertyName === 'CenterVertically') {
            return $this->_model->PrintOptions->VerticalCentered;
        }
        if ($propertyName === 'HorizontalPageBreaks') {
            return self::GetOrderedBreaks($this->_model->HorizontalPageBreaks);
        }
        if ($propertyName === 'VerticalPageBreaks') {
            return self::GetOrderedBreaks($this->_model->VerticalPageBreaks);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'PaperSize') {
            $this->_model->PaperSize = $value;
            return;
        }
        if ($propertyName === 'Orientation') {
            switch ($value) {
                case PageOrientationType::Portrait:
                    $this->_model->Orientation = PageOrientation::Portrait;
                    break;
                case PageOrientationType::Landscape:
                    $this->_model->Orientation = PageOrientation::Landscape;
                    break;
                default:
                    $this->_model->Orientation = PageOrientation::Default;
                    break;
            }
            return;
        }
        if ($propertyName === 'FirstPageNumber') {
            if ($value !== null && $value <= 0) {
                throw new CellsException('FirstPageNumber must be positive.');
            }
            $this->_model->FirstPageNumber = $value;
            return;
        }
        if ($propertyName === 'Scale') {
            if ($value !== null && ($value < 10 || $value > 400)) {
                throw new CellsException('Scale must be between 10 and 400.');
            }
            $this->_model->Scale = $value;
            return;
        }
        if ($propertyName === 'FitToPagesWide') {
            if ($value !== null && $value < 0) {
                throw new CellsException('FitToPagesWide must be zero or greater.');
            }
            $this->_model->FitToWidth = $value;
            return;
        }
        if ($propertyName === 'FitToPagesTall') {
            if ($value !== null && $value < 0) {
                throw new CellsException('FitToPagesTall must be zero or greater.');
            }
            $this->_model->FitToHeight = $value;
            return;
        }
        if ($propertyName === 'PrintArea') {
            $this->_model->PrintArea = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'PrintTitleRows') {
            $this->_model->PrintTitleRows = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'PrintTitleColumns') {
            $this->_model->PrintTitleColumns = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'LeftMargin') {
            $this->_model->Margins->Left = self::ValidateMargin(self::ToInches($value), 'LeftMargin');
            return;
        }
        if ($propertyName === 'RightMargin') {
            $this->_model->Margins->Right = self::ValidateMargin(self::ToInches($value), 'RightMargin');
            return;
        }
        if ($propertyName === 'TopMargin') {
            $this->_model->Margins->Top = self::ValidateMargin(self::ToInches($value), 'TopMargin');
            return;
        }
        if ($propertyName === 'BottomMargin') {
            $this->_model->Margins->Bottom = self::ValidateMargin(self::ToInches($value), 'BottomMargin');
            return;
        }
        if ($propertyName === 'HeaderMargin') {
            $this->_model->Margins->Header = self::ValidateMargin(self::ToInches($value), 'HeaderMargin');
            return;
        }
        if ($propertyName === 'FooterMargin') {
            $this->_model->Margins->Footer = self::ValidateMargin(self::ToInches($value), 'FooterMargin');
            return;
        }
        if ($propertyName === 'LeftMarginInch') {
            $this->_model->Margins->Left = self::ValidateMargin($value, 'LeftMarginInch');
            return;
        }
        if ($propertyName === 'RightMarginInch') {
            $this->_model->Margins->Right = self::ValidateMargin($value, 'RightMarginInch');
            return;
        }
        if ($propertyName === 'TopMarginInch') {
            $this->_model->Margins->Top = self::ValidateMargin($value, 'TopMarginInch');
            return;
        }
        if ($propertyName === 'BottomMarginInch') {
            $this->_model->Margins->Bottom = self::ValidateMargin($value, 'BottomMarginInch');
            return;
        }
        if ($propertyName === 'HeaderMarginInch') {
            $this->_model->Margins->Header = self::ValidateMargin($value, 'HeaderMarginInch');
            return;
        }
        if ($propertyName === 'FooterMarginInch') {
            $this->_model->Margins->Footer = self::ValidateMargin($value, 'FooterMarginInch');
            return;
        }
        if ($propertyName === 'LeftHeader') {
            $this->_model->HeaderFooter->LeftHeader = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'CenterHeader') {
            $this->_model->HeaderFooter->CenterHeader = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'RightHeader') {
            $this->_model->HeaderFooter->RightHeader = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'LeftFooter') {
            $this->_model->HeaderFooter->LeftFooter = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'CenterFooter') {
            $this->_model->HeaderFooter->CenterFooter = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'RightFooter') {
            $this->_model->HeaderFooter->RightFooter = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'PrintGridlines') {
            $this->_model->PrintOptions->GridLines = $value;
            return;
        }
        if ($propertyName === 'PrintHeadings') {
            $this->_model->PrintOptions->Headings = $value;
            return;
        }
        if ($propertyName === 'CenterHorizontally') {
            $this->_model->PrintOptions->HorizontalCentered = $value;
            return;
        }
        if ($propertyName === 'CenterVertically') {
            $this->_model->PrintOptions->VerticalCentered = $value;
            return;
        }
        if ($propertyName === 'HorizontalPageBreaks') {
            throw new \RuntimeException('HorizontalPageBreaks is read-only.');
        }
        if ($propertyName === 'VerticalPageBreaks') {
            throw new \RuntimeException('VerticalPageBreaks is read-only.');
        }
    }

    public function AddHorizontalPageBreak($rowIndex)
    {
        if ($rowIndex < 0) {
            throw new CellsException('Horizontal page break row index must be non-negative.');
        }
        self::AddDistinct($this->_model->HorizontalPageBreaks, $rowIndex);
    }

    public function AddVerticalPageBreak($columnIndex)
    {
        if ($columnIndex < 0) {
            throw new CellsException('Vertical page break column index must be non-negative.');
        }
        self::AddDistinct($this->_model->VerticalPageBreaks, $columnIndex);
    }

    public function ClearHorizontalPageBreaks()
    {
        $this->_model->HorizontalPageBreaks = array();
    }

    public function ClearVerticalPageBreaks()
    {
        $this->_model->VerticalPageBreaks = array();
    }

    private static function GetOrderedBreaks($breaks)
    {
        $orderedBreaks = $breaks;
        sort($orderedBreaks);
        return $orderedBreaks;
    }

    private static function AddDistinct(&$collection, $value)
    {
        if (!in_array($value, $collection, true)) {
            $collection[] = $value;
        }
    }

    private static function ValidateMargin($value, $propertyName)
    {
        if ($value < 0) {
            throw new CellsException($propertyName . ' must be zero or greater.');
        }
        return $value;
    }

    private static function ToCentimeters($inches)
    {
        return $inches * self::CentimetersPerInch;
    }

    private static function ToInches($centimeters)
    {
        return $centimeters / self::CentimetersPerInch;
    }

    private static function NormalizeText($value)
    {
        if ($value === null || trim($value) === '') {
            return null;
        }
        return trim($value);
    }
}