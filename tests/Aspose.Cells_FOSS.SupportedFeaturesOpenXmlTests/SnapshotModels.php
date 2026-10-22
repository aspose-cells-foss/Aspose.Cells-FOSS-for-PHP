<?php

namespace Aspose\Cells_FOSS\SupportedFeaturesOpenXmlTests;

final class OpenXmlDateSystem
{
    const Windows1900 = 0;
    const Mac1904 = 1;
}

final class OpenXmlReadContext
{
    public function __construct($package = null, $workbookPath = null, $workbookDocument = null, $workbookRoot = null, $workbookRelationships = array(), $sheets = array(), $sharedStrings = array(), $dateSystem = OpenXmlDateSystem::Windows1900, $styles = null)
    {
        $this->Package = $package;
        $this->WorkbookPath = $workbookPath;
        $this->WorkbookDocument = $workbookDocument;
        $this->WorkbookRoot = $workbookRoot;
        $this->WorkbookRelationships = $workbookRelationships;
        $this->Sheets = $sheets;
        $this->SharedStrings = $sharedStrings;
        $this->DateSystem = $dateSystem;
        $this->Styles = $styles;
    }

    public $Package;
    public $WorkbookPath;
    public $WorkbookDocument;
    public $WorkbookRoot;
    public $WorkbookRelationships = array();
    public $Sheets = array();
    public $SharedStrings = array();
    public $DateSystem;
    public $Styles;
}

final class OpenXmlCellDescriptor
{
    public $SheetName = '';
    public $CellReference = '';
    public $Cell;

    public function __construct($sheetName, $cellReference, $cell)
    {
        $this->SheetName = $sheetName;
        $this->CellReference = $cellReference;
        $this->Cell = $cell;
    }
}

final class OpenXmlFontDescriptor
{
    public $Name;
    public $Size;
    public $Bold;
    public $Italic;
    public $Underline;
    public $StrikeThrough;
    public $Color;

    public function __construct($name, $size, $bold, $italic, $underline, $strikeThrough, $color)
    {
        $this->Name = $name;
        $this->Size = $size;
        $this->Bold = $bold;
        $this->Italic = $italic;
        $this->Underline = $underline;
        $this->StrikeThrough = $strikeThrough;
        $this->Color = $color;
    }

    private static $default = null;

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlFontDescriptor('Calibri', ComparisonValueHelpers::NormalizeDouble(11.0), false, false, false, false, ComparisonValueHelpers::DefaultColorHex);
        }

        return self::$default;
    }
}

final class OpenXmlFillDescriptor
{
    public $Pattern;
    public $ForegroundColor;
    public $BackgroundColor;

    public function __construct($pattern, $foregroundColor, $backgroundColor)
    {
        $this->Pattern = $pattern;
        $this->ForegroundColor = $foregroundColor;
        $this->BackgroundColor = $backgroundColor;
    }

    private static $default = null;

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlFillDescriptor('None', ComparisonValueHelpers::DefaultColorHex, ComparisonValueHelpers::DefaultColorHex);
        }

        return self::$default;
    }
}

final class OpenXmlBorderSideDescriptor
{
    public $Style;
    public $Color;

    public function __construct($style, $color)
    {
        $this->Style = $style;
        $this->Color = $color;
    }

    private static $default = null;

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlBorderSideDescriptor('None', ComparisonValueHelpers::DefaultColorHex);
        }

        return self::$default;
    }
}

final class OpenXmlBorderDescriptor
{
    public $Left;
    public $Right;
    public $Top;
    public $Bottom;
    public $Diagonal;
    public $DiagonalUp;
    public $DiagonalDown;

    public function __construct($left, $right, $top, $bottom, $diagonal, $diagonalUp, $diagonalDown)
    {
        $this->Left = $left;
        $this->Right = $right;
        $this->Top = $top;
        $this->Bottom = $bottom;
        $this->Diagonal = $diagonal;
        $this->DiagonalUp = $diagonalUp;
        $this->DiagonalDown = $diagonalDown;
    }

    private static $default = null;

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlBorderDescriptor(OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), false, false);
        }

        return self::$default;
    }
}

final class CellSnapshot
{
    public $SheetName;
    public $CellReference;
    public $CellType;
    public $Value;
    public $Formula;

    public function __construct($sheetName, $cellReference, $cellType, $value, $formula)
    {
        $this->SheetName = $sheetName;
        $this->CellReference = $cellReference;
        $this->CellType = $cellType;
        $this->Value = $value;
        $this->Formula = $formula;
    }

    public function ToString()
    {
        return 'Sheet=' . $this->SheetName . '; Cell=' . $this->CellReference . '; Type=' . $this->CellType . '; Formula=' . $this->Formula . '; Value=' . $this->Value;
    }

    public function __toString()
    {
        return $this->ToString();
    }
}

final class StyleSnapshot
{
    public $SheetName = '';
    public $CellReference = '';
    public $FontName = 'Calibri';
    public $FontSize = '11';
    public $FontBold = false;
    public $FontItalic = false;
    public $FontUnderline = false;
    public $FontStrikeThrough = false;
    public $FontColor = ComparisonValueHelpers::DefaultColorHex;
    public $FillPattern = 'None';
    public $FillForegroundColor = ComparisonValueHelpers::DefaultColorHex;
    public $FillBackgroundColor = ComparisonValueHelpers::DefaultColorHex;
    public $LeftBorderStyle = 'None';
    public $LeftBorderColor = ComparisonValueHelpers::DefaultColorHex;
    public $RightBorderStyle = 'None';
    public $RightBorderColor = ComparisonValueHelpers::DefaultColorHex;
    public $TopBorderStyle = 'None';
    public $TopBorderColor = ComparisonValueHelpers::DefaultColorHex;
    public $BottomBorderStyle = 'None';
    public $BottomBorderColor = ComparisonValueHelpers::DefaultColorHex;
    public $DiagonalBorderStyle = 'None';
    public $DiagonalBorderColor = ComparisonValueHelpers::DefaultColorHex;
    public $DiagonalUp = false;
    public $DiagonalDown = false;
    public $HorizontalAlignment = 'General';
    public $VerticalAlignment = 'Bottom';
    public $WrapText = false;
    public $IndentLevel = 0;
    public $TextRotation = 0;
    public $ShrinkToFit = false;
    public $ReadingOrder = 0;
    public $RelativeIndent = 0;
    public $NumberFormatId = 0;
    public $NumberFormatCode = '';
    public $IsLocked = true;
    public $IsHidden = false;

    public function __construct(
        $sheetName,
        $cellReference,
        $fontName,
        $fontSize,
        $fontBold,
        $fontItalic,
        $fontUnderline,
        $fontStrikeThrough,
        $fontColor,
        $fillPattern,
        $fillForegroundColor,
        $fillBackgroundColor,
        $leftBorderStyle,
        $leftBorderColor,
        $rightBorderStyle,
        $rightBorderColor,
        $topBorderStyle,
        $topBorderColor,
        $bottomBorderStyle,
        $bottomBorderColor,
        $diagonalBorderStyle,
        $diagonalBorderColor,
        $diagonalUp,
        $diagonalDown,
        $horizontalAlignment,
        $verticalAlignment,
        $wrapText,
        $indentLevel,
        $textRotation,
        $shrinkToFit,
        $readingOrder,
        $relativeIndent,
        $numberFormatId,
        $numberFormatCode,
        $isLocked,
        $isHidden
    ) {
        $this->SheetName = $sheetName;
        $this->CellReference = $cellReference;
        $this->FontName = $fontName;
        $this->FontSize = $fontSize;
        $this->FontBold = $fontBold;
        $this->FontItalic = $fontItalic;
        $this->FontUnderline = $fontUnderline;
        $this->FontStrikeThrough = $fontStrikeThrough;
        $this->FontColor = $fontColor;
        $this->FillPattern = $fillPattern;
        $this->FillForegroundColor = $fillForegroundColor;
        $this->FillBackgroundColor = $fillBackgroundColor;
        $this->LeftBorderStyle = $leftBorderStyle;
        $this->LeftBorderColor = $leftBorderColor;
        $this->RightBorderStyle = $rightBorderStyle;
        $this->RightBorderColor = $rightBorderColor;
        $this->TopBorderStyle = $topBorderStyle;
        $this->TopBorderColor = $topBorderColor;
        $this->BottomBorderStyle = $bottomBorderStyle;
        $this->BottomBorderColor = $bottomBorderColor;
        $this->DiagonalBorderStyle = $diagonalBorderStyle;
        $this->DiagonalBorderColor = $diagonalBorderColor;
        $this->DiagonalUp = $diagonalUp;
        $this->DiagonalDown = $diagonalDown;
        $this->HorizontalAlignment = $horizontalAlignment;
        $this->VerticalAlignment = $verticalAlignment;
        $this->WrapText = $wrapText;
        $this->IndentLevel = $indentLevel;
        $this->TextRotation = $textRotation;
        $this->ShrinkToFit = $shrinkToFit;
        $this->ReadingOrder = $readingOrder;
        $this->RelativeIndent = $relativeIndent;
        $this->NumberFormatId = $numberFormatId;
        $this->NumberFormatCode = $numberFormatCode;
        $this->IsLocked = $isLocked;
        $this->IsHidden = $isHidden;
    }

    public function CopyWithLocation($sheetName, $cellReference)
    {
        $clone = clone $this;
        $clone->SheetName = $sheetName;
        $clone->CellReference = $cellReference;
        return $clone;
    }

    private static $default = null;

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new StyleSnapshot(
                '',
                '',
                'Calibri',
                ComparisonValueHelpers::NormalizeDouble(11.0),
                false,
                false,
                false,
                false,
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                false,
                false,
                'General',
                'Bottom',
                false,
                0,
                0,
                false,
                0,
                0,
                0,
                '',
                true,
                false
            );
        }

        return self::$default;
    }
}

final class OpenXmlStylesContext
{
    public $CellFormats = array();
    public $DateStyleIndexes = array();

    public function __construct($cellFormats, $dateStyleIndexes)
    {
        $this->CellFormats = $cellFormats;
        $this->DateStyleIndexes = $dateStyleIndexes;
    }

    private static $default = null;

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlStylesContext(array(StyleSnapshot::Default()), array());
        }

        return self::$default;
    }
}

final class WorkbookSnapshot
{
    public $UsesDate1904 = false;
    public $HasSharedStrings = false;
    public $Sheets = array();
    public $Cells = array();
    public $Styles = array();
    public $Worksheets = array();
}

final class SheetSnapshot
{
    public $Name = '';
    public $Visibility = 'Visible';

    public function __construct($name, $visibility)
    {
        $this->Name = $name;
        $this->Visibility = $visibility;
    }
}

final class DefinedNameSnapshot
{
    public $Name = '';
    public $Formula = '';
    public $LocalSheetIndex = null;
    public $Hidden = false;
    public $Comment = '';

    public function __construct($name, $formula, $localSheetIndex, $hidden, $comment)
    {
        $this->Name = $name;
        $this->Formula = $formula;
        $this->LocalSheetIndex = $localSheetIndex;
        $this->Hidden = $hidden;
        $this->Comment = $comment;
    }
}

final class WorksheetSnapshot
{
    public $Name = '';
    public $Visibility = 'Visible';
    public $Dimension = '';
    public $Rows = array();
    public $Columns = array();
    public $MergeReferences = array();
    public $Hyperlinks = array();
    public $Validations = array();
    public $ConditionalFormattings = array();
    public $PageSetup = null;

    public function __construct($name, $visibility, $dimension, $rows, $columns, $mergeReferences, $hyperlinks, $validations, $conditionalFormattings, $pageSetup)
    {
        $this->Name = $name;
        $this->Visibility = $visibility;
        $this->Dimension = $dimension;
        $this->Rows = $rows;
        $this->Columns = $columns;
        $this->MergeReferences = $mergeReferences;
        $this->Hyperlinks = $hyperlinks;
        $this->Validations = $validations;
        $this->ConditionalFormattings = $conditionalFormattings;
        $this->PageSetup = $pageSetup;
    }
}

final class RowSnapshot
{
    public $Height = null;
    public $Hidden = false;

    public function __construct($height, $hidden)
    {
        $this->Height = $height;
        $this->Hidden = $hidden;
    }
}

final class ColumnSnapshot
{
    public $MinColumnIndex = 0;
    public $MaxColumnIndex = 0;
    public $Width = null;
    public $Hidden = false;

    public function __construct($minColumnIndex, $maxColumnIndex, $width, $hidden)
    {
        $this->MinColumnIndex = $minColumnIndex;
        $this->MaxColumnIndex = $maxColumnIndex;
        $this->Width = $width;
        $this->Hidden = $hidden;
    }
}

final class HyperlinkSnapshot
{
    public $Area = '';
    public $Address = '';
    public $Location = '';
    public $ScreenTip = '';
    public $Display = '';

    public function __construct($area, $address, $location, $screenTip, $display)
    {
        $this->Area = $area;
        $this->Address = $address;
        $this->Location = $location;
        $this->ScreenTip = $screenTip;
        $this->Display = $display;
    }
}

final class ValidationSnapshot
{
    public $Sqref = '';
    public $Type = '';
    public $Operator = '';
    public $ErrorStyle = 'stop';
    public $AllowBlank = false;
    public $ShowDropDown = false;
    public $ShowInputMessage = false;
    public $ShowErrorMessage = false;
    public $PromptTitle = '';
    public $Prompt = '';
    public $ErrorTitle = '';
    public $Error = '';
    public $Formula1 = '';
    public $Formula2 = '';

    public function __construct($sqref, $type, $operator, $errorStyle, $allowBlank, $showDropDown, $showInputMessage, $showErrorMessage, $promptTitle, $prompt, $errorTitle, $error, $formula1, $formula2)
    {
        $this->Sqref = $sqref;
        $this->Type = $type;
        $this->Operator = $operator;
        $this->ErrorStyle = $errorStyle;
        $this->AllowBlank = $allowBlank;
        $this->ShowDropDown = $showDropDown;
        $this->ShowInputMessage = $showInputMessage;
        $this->ShowErrorMessage = $showErrorMessage;
        $this->PromptTitle = $promptTitle;
        $this->Prompt = $prompt;
        $this->ErrorTitle = $errorTitle;
        $this->Error = $error;
        $this->Formula1 = $formula1;
        $this->Formula2 = $formula2;
    }
}

final class ConditionalFormattingSnapshot
{
    public $Sqref = '';
    public $Rules = array();

    public function __construct($sqref, $rules)
    {
        $this->Sqref = $sqref;
        $this->Rules = $rules;
    }
}

final class ConditionalFormattingRuleSnapshot
{
    public $Type = '';
    public $Operator = '';
    public $Priority = '';
    public $StopIfTrue = '';
    public $DxfId = '';
    public $Formulas = array();

    public function __construct($type, $operator, $priority, $stopIfTrue, $dxfId, $formulas)
    {
        $this->Type = $type;
        $this->Operator = $operator;
        $this->Priority = $priority;
        $this->StopIfTrue = $stopIfTrue;
        $this->DxfId = $dxfId;
        $this->Formulas = $formulas;
    }
}

final class PageSetupSnapshot
{
    public $PrintArea = '';
    public $PrintTitleRows = '';
    public $PrintTitleColumns = '';
    public $LeftMargin = null;
    public $RightMargin = null;
    public $TopMargin = null;
    public $BottomMargin = null;
    public $HeaderMargin = null;
    public $FooterMargin = null;
    public $PaperSize = null;
    public $Orientation = '';
    public $FirstPageNumber = null;
    public $Scale = null;
    public $FitToWidth = null;
    public $FitToHeight = null;
    public $PrintGridLines = false;
    public $PrintHeadings = false;
    public $HorizontalCentered = false;
    public $VerticalCentered = false;
    public $OddHeader = '';
    public $OddFooter = '';
    public $RowBreaks = array();
    public $ColumnBreaks = array();

    public function __construct($printArea, $printTitleRows, $printTitleColumns, $leftMargin, $rightMargin, $topMargin, $bottomMargin, $headerMargin, $footerMargin, $paperSize, $orientation, $firstPageNumber, $scale, $fitToWidth, $fitToHeight, $printGridLines, $printHeadings, $horizontalCentered, $verticalCentered, $oddHeader, $oddFooter, $rowBreaks, $columnBreaks)
    {
        $this->PrintArea = $printArea;
        $this->PrintTitleRows = $printTitleRows;
        $this->PrintTitleColumns = $printTitleColumns;
        $this->LeftMargin = $leftMargin;
        $this->RightMargin = $rightMargin;
        $this->TopMargin = $topMargin;
        $this->BottomMargin = $bottomMargin;
        $this->HeaderMargin = $headerMargin;
        $this->FooterMargin = $footerMargin;
        $this->PaperSize = $paperSize;
        $this->Orientation = $orientation;
        $this->FirstPageNumber = $firstPageNumber;
        $this->Scale = $scale;
        $this->FitToWidth = $fitToWidth;
        $this->FitToHeight = $fitToHeight;
        $this->PrintGridLines = $printGridLines;
        $this->PrintHeadings = $printHeadings;
        $this->HorizontalCentered = $horizontalCentered;
        $this->VerticalCentered = $verticalCentered;
        $this->OddHeader = $oddHeader;
        $this->OddFooter = $oddFooter;
        $this->RowBreaks = $rowBreaks;
        $this->ColumnBreaks = $columnBreaks;
    }
}
