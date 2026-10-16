<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ChartCompanionFile;
use Aspose\Cells_FOSS\Core\ChartModel;
use Aspose\Cells_FOSS\Core\PictureModel;
use Aspose\Cells_FOSS\Core\ShapeImageModel;
use Aspose\Cells_FOSS\Core\ShapeModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

// Loads worksheet drawing parts (pictures, shapes, charts) from the package.
final class XlsxWorkbookPicturesLoader
{
    private function __construct()
    {
    }

    public static function LoadPictures(WorksheetModel $worksheetModel, XElement $worksheetRoot, ZipArchive $archive, $worksheetUri, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->Pictures = array();
        $worksheetModel->ShapeImages = array();
        $worksheetModel->Shapes = array();
        $worksheetModel->Charts = array();

        $drawingUri = self::FindDrawingUri($archive, $worksheetUri);
        if ($drawingUri === null || $drawingUri === '') {
            return;
        }

        $drawingEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $drawingUri);
        if ($drawingEntry === null) {
            $issue = new LoadIssue('PIC-R001', DiagnosticSeverity::Recoverable, "Drawing part '" . $drawingUri . "' was referenced but not found; pictures were skipped.", true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return;
        }

        $drawingDocument = XlsxWorkbookArchiveHelpers::LoadDocument($drawingEntry);
        $drawingRoot = $drawingDocument->Root;
        if ($drawingRoot === null) {
            return;
        }

        $imageTargets = self::LoadDrawingImageTargets($archive, $drawingUri);
        $chartTargets = self::LoadDrawingChartTargets($archive, $drawingUri);

        self::LoadTwoCellAnchorPictures($worksheetModel, $drawingRoot, $imageTargets, $archive, $diagnostics, $options, $sheetName);
        self::LoadOneCellAnchorPictures($worksheetModel, $drawingRoot, $imageTargets, $archive, $diagnostics, $options, $sheetName);
        self::LoadShapeImages($worksheetModel, $imageTargets, $archive);
        self::LoadTwoCellAnchorShapes($worksheetModel, $drawingRoot);
        self::LoadOneCellAnchorShapes($worksheetModel, $drawingRoot);
        self::LoadTwoCellAnchorCharts($worksheetModel, $drawingRoot, $chartTargets, $archive, $diagnostics, $options, $sheetName);
        self::LoadOneCellAnchorCharts($worksheetModel, $drawingRoot, $chartTargets, $archive, $diagnostics, $options, $sheetName);
    }

    private static function LoadTwoCellAnchorPictures(WorksheetModel $worksheetModel, XElement $drawingRoot, $imageTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        foreach ($drawingRoot->Elements(XlsxWorkbookPictures::XdrNs . 'twoCellAnchor') as $anchor) {
            $pic = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'pic');
            if ($pic === null) {
                continue;
            }

            $fromEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'from');
            $toEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'to');
            if ($fromEl === null || $toEl === null) {
                continue;
            }

            $model = self::ParsePicture($pic, $fromEl, $toEl, $imageTargets, $archive, $diagnostics, $options, $sheetName);
            if ($model !== null) {
                $worksheetModel->Pictures[] = $model;
            }
        }
    }

    private static function LoadOneCellAnchorPictures(WorksheetModel $worksheetModel, XElement $drawingRoot, $imageTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        foreach ($drawingRoot->Elements(XlsxWorkbookPictures::XdrNs . 'oneCellAnchor') as $anchor) {
            $pic = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'pic');
            if ($pic === null) {
                continue;
            }

            $fromEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'from');
            if ($fromEl === null) {
                continue;
            }

            $extEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'ext');
            $model = self::ParseOneCellPicture($pic, $fromEl, $extEl, $imageTargets, $archive, $diagnostics, $options, $sheetName);
            if ($model !== null) {
                $worksheetModel->Pictures[] = $model;
            }
        }
    }

    private static function LoadShapeImages(WorksheetModel $worksheetModel, $imageTargets, ZipArchive $archive)
    {
        $worksheetModel->ShapeImages = array();

        // Collect rIds already claimed by loaded top-level pictures
        $claimedRIds = array();
        foreach ($worksheetModel->Pictures as $pic) {
            if (!($pic->OriginalRId === null || $pic->OriginalRId === '')) {
                $claimedRIds[$pic->OriginalRId] = true;
            }
        }

        // Load every image relationship that was NOT claimed by a top-level picture.
        // These images are referenced from within shape/group raw XML and must be preserved
        // so their r:embed rId references remain valid after save.
        foreach ($imageTargets as $rId => $targetUri) {
            if (isset($claimedRIds[$rId])) {
                continue;
            }

            $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $targetUri);
            if ($entry === null) {
                continue;
            }

            $imageData = $entry->getContent();

            if (strlen($imageData) === 0) {
                continue;
            }

            $shapeImage = new ShapeImageModel();
            $shapeImage->OriginalRId = $rId;
            $shapeImage->Extension = Picture::DetectExtension($imageData);
            $shapeImage->ImageData = $imageData;
            $worksheetModel->ShapeImages[] = $shapeImage;
        }
    }

    private static function LoadTwoCellAnchorShapes(WorksheetModel $worksheetModel, XElement $drawingRoot)
    {
        foreach ($drawingRoot->Elements(XlsxWorkbookPictures::XdrNs . 'twoCellAnchor') as $anchor) {
            $fromEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'from');
            $toEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'to');
            if ($fromEl === null || $toEl === null) {
                continue;
            }

            $sp = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'sp');
            if ($sp !== null) {
                $model = self::ParseShape($sp, $fromEl, $toEl);
                if ($model !== null) {
                    $worksheetModel->Shapes[] = $model;
                }
                continue;
            }

            $grpSp = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'grpSp');
            if ($grpSp !== null) {
                $model = self::ParseGroupShape($grpSp, $fromEl, $toEl);
                if ($model !== null) {
                    $worksheetModel->Shapes[] = $model;
                }
                continue;
            }

            $cxnSp = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'cxnSp');
            if ($cxnSp !== null) {
                $model = self::ParseConnector($cxnSp, $fromEl, $toEl);
                if ($model !== null) {
                    $worksheetModel->Shapes[] = $model;
                }
            }
        }
    }

    private static function LoadOneCellAnchorShapes(WorksheetModel $worksheetModel, XElement $drawingRoot)
    {
        foreach ($drawingRoot->Elements(XlsxWorkbookPictures::XdrNs . 'oneCellAnchor') as $anchor) {
            $fromEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'from');
            if ($fromEl === null) {
                continue;
            }

            $extEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'ext');

            $sp = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'sp');
            if ($sp !== null) {
                $model = self::ParseOneCellShape($sp, $fromEl, $extEl);
                if ($model !== null) {
                    $worksheetModel->Shapes[] = $model;
                }
                continue;
            }

            $grpSp = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'grpSp');
            if ($grpSp !== null) {
                $model = self::ParseOneCellGroupShape($grpSp, $fromEl, $extEl);
                if ($model !== null) {
                    $worksheetModel->Shapes[] = $model;
                }
                continue;
            }

            $cxnSp = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'cxnSp');
            if ($cxnSp !== null) {
                $model = self::ParseOneCellConnector($cxnSp, $fromEl, $extEl);
                if ($model !== null) {
                    $worksheetModel->Shapes[] = $model;
                }
            }
        }
    }

    private static function ParseShape(XElement $sp, XElement $fromEl, XElement $toEl)
    {
        $model = new ShapeModel();
        $model->Name = self::GetShapeName($sp);
        $model->UpperLeftColumn = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->LowerRightColumnOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->LowerRightRow = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->LowerRightRowOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->GeometryType = self::GetGeometryType($sp);
        $model->RawStyleXml = self::GetRawXml($sp->Element(XlsxWorkbookPictures::XdrNs . 'style'));
        $model->RawTxBodyXml = self::GetRawXml($sp->Element(XlsxWorkbookPictures::XdrNs . 'txBody'));
        $model->RawElementXml = $sp->toString();
        self::LoadShapeSpPrExtents($sp, $model);
        return $model;
    }

    private static function ParseOneCellShape(XElement $sp, XElement $fromEl, $extEl)
    {
        $fromCol = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $fromRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $cx = 0;
        $cy = 0;

        if ($extEl !== null) {
            $cxText = (string) $extEl->Attribute('cx');
            if ($cxText === '') {
                $cxText = '0';
            }
            if (is_numeric($cxText)) {
                $cx = (int) $cxText;
            }

            $cyText = (string) $extEl->Attribute('cy');
            if ($cyText === '') {
                $cyText = '0';
            }
            if (is_numeric($cyText)) {
                $cy = (int) $cyText;
            }
        }

        $colSpan = $cx > 0 ? (int) ($cx / 609600) + 1 : 1;
        $rowSpan = $cy > 0 ? (int) ($cy / 190500) + 1 : 1;

        $model = new ShapeModel();
        $model->Name = self::GetShapeName($sp);
        $model->UpperLeftColumn = $fromCol;
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = $fromRow;
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = $fromCol + $colSpan;
        $model->LowerRightRow = $fromRow + $rowSpan;
        $model->ExtentCx = $cx;
        $model->ExtentCy = $cy;
        $model->GeometryType = self::GetGeometryType($sp);
        $model->RawStyleXml = self::GetRawXml($sp->Element(XlsxWorkbookPictures::XdrNs . 'style'));
        $model->RawTxBodyXml = self::GetRawXml($sp->Element(XlsxWorkbookPictures::XdrNs . 'txBody'));
        $model->RawElementXml = $sp->toString();
        return $model;
    }

    private static function ParseConnector(XElement $cxnSp, XElement $fromEl, XElement $toEl)
    {
        $model = new ShapeModel();
        $model->Name = self::GetConnectorName($cxnSp);
        $model->UpperLeftColumn = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->LowerRightColumnOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->LowerRightRow = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->LowerRightRowOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->GeometryType = self::GetGeometryType($cxnSp);
        $model->RawElementXml = $cxnSp->toString();
        return $model;
    }

    private static function ParseOneCellConnector(XElement $cxnSp, XElement $fromEl, $extEl)
    {
        $fromCol = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $fromRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $cx = 0;
        $cy = 0;

        if ($extEl !== null) {
            $cxText = (string) $extEl->Attribute('cx');
            if ($cxText === '') {
                $cxText = '0';
            }
            if (is_numeric($cxText)) {
                $cx = (int) $cxText;
            }

            $cyText = (string) $extEl->Attribute('cy');
            if ($cyText === '') {
                $cyText = '0';
            }
            if (is_numeric($cyText)) {
                $cy = (int) $cyText;
            }
        }

        $colSpan = $cx > 0 ? (int) ($cx / 609600) + 1 : 1;
        $rowSpan = $cy > 0 ? (int) ($cy / 190500) + 1 : 1;

        $model = new ShapeModel();
        $model->Name = self::GetConnectorName($cxnSp);
        $model->UpperLeftColumn = $fromCol;
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = $fromRow;
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = $fromCol + $colSpan;
        $model->LowerRightRow = $fromRow + $rowSpan;
        $model->ExtentCx = $cx;
        $model->ExtentCy = $cy;
        $model->GeometryType = self::GetGeometryType($cxnSp);
        $model->RawElementXml = $cxnSp->toString();
        return $model;
    }

    private static function ParseGroupShape(XElement $grpSp, XElement $fromEl, XElement $toEl)
    {
        $model = new ShapeModel();
        $model->Name = self::GetGroupShapeName($grpSp);
        $model->UpperLeftColumn = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->LowerRightColumnOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->LowerRightRow = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->LowerRightRowOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->RawElementXml = $grpSp->toString();
        return $model;
    }

    private static function ParseOneCellGroupShape(XElement $grpSp, XElement $fromEl, $extEl)
    {
        $fromCol = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $fromRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $cx = 0;
        $cy = 0;

        if ($extEl !== null) {
            $cxText = (string) $extEl->Attribute('cx');
            if ($cxText === '') {
                $cxText = '0';
            }
            if (is_numeric($cxText)) {
                $cx = (int) $cxText;
            }

            $cyText = (string) $extEl->Attribute('cy');
            if ($cyText === '') {
                $cyText = '0';
            }
            if (is_numeric($cyText)) {
                $cy = (int) $cyText;
            }
        }

        $colSpan = $cx > 0 ? (int) ($cx / 609600) + 1 : 1;
        $rowSpan = $cy > 0 ? (int) ($cy / 190500) + 1 : 1;

        $model = new ShapeModel();
        $model->Name = self::GetGroupShapeName($grpSp);
        $model->UpperLeftColumn = $fromCol;
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = $fromRow;
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = $fromCol + $colSpan;
        $model->LowerRightRow = $fromRow + $rowSpan;
        $model->ExtentCx = $cx;
        $model->ExtentCy = $cy;
        $model->RawElementXml = $grpSp->toString();
        return $model;
    }

    private static function GetGroupShapeName(XElement $grpSp)
    {
        $nvGrpSpPr = $grpSp->Element(XlsxWorkbookPictures::XdrNs . 'nvGrpSpPr');
        if ($nvGrpSpPr === null) {
            return '';
        }

        $cNvPr = $nvGrpSpPr->Element(XlsxWorkbookPictures::XdrNs . 'cNvPr');
        return $cNvPr === null ? '' : (string) $cNvPr->Attribute('name');
    }

    private static function GetConnectorName(XElement $cxnSp)
    {
        $nvCxnSpPr = $cxnSp->Element(XlsxWorkbookPictures::XdrNs . 'nvCxnSpPr');
        if ($nvCxnSpPr === null) {
            return '';
        }

        $cNvPr = $nvCxnSpPr->Element(XlsxWorkbookPictures::XdrNs . 'cNvPr');
        return $cNvPr === null ? '' : (string) $cNvPr->Attribute('name');
    }

    private static function GetShapeName(XElement $sp)
    {
        $nvSpPr = $sp->Element(XlsxWorkbookPictures::XdrNs . 'nvSpPr');
        if ($nvSpPr === null) {
            return '';
        }

        $cNvPr = $nvSpPr->Element(XlsxWorkbookPictures::XdrNs . 'cNvPr');
        return $cNvPr === null ? '' : (string) $cNvPr->Attribute('name');
    }

    private static function GetGeometryType(XElement $sp)
    {
        $spPr = $sp->Element(XlsxWorkbookPictures::XdrNs . 'spPr');
        if ($spPr === null) {
            return 'rect';
        }

        $prstGeom = $spPr->Element(XlsxWorkbookPictures::ANs . 'prstGeom');
        return $prstGeom === null ? 'rect' : (string) $prstGeom->Attribute('prst');
    }

    private static function GetRawXml($element)
    {
        return $element === null ? null : $element->toString();
    }

    private static function LoadShapeSpPrExtents(XElement $sp, ShapeModel $model)
    {
        $spPr = $sp->Element(XlsxWorkbookPictures::XdrNs . 'spPr');
        if ($spPr === null) {
            return;
        }

        $xfrm = $spPr->Element(XlsxWorkbookPictures::ANs . 'xfrm');
        if ($xfrm === null) {
            return;
        }

        $ext = $xfrm->Element(XlsxWorkbookPictures::ANs . 'ext');
        if ($ext === null) {
            return;
        }

        $cx = 0;
        $cy = 0;
        $cxText = (string) $ext->Attribute('cx');
        if ($cxText === '') {
            $cxText = '0';
        }
        if (is_numeric($cxText)) {
            $cx = (int) $cxText;
        }

        $cyText = (string) $ext->Attribute('cy');
        if ($cyText === '') {
            $cyText = '0';
        }
        if (is_numeric($cyText)) {
            $cy = (int) $cyText;
        }

        $model->ExtentCx = $cx;
        $model->ExtentCy = $cy;
    }

    private static function ParsePicture(XElement $pic, XElement $fromEl, XElement $toEl, $imageTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $rId = self::GetBlipRId($pic);
        if ($rId === null || $rId === '') {
            return null;
        }

        $name = self::GetPictureName($pic);
        $imageData = self::LoadImageData($rId, $imageTargets, $archive, $diagnostics, $options, $sheetName, $name);
        if ($imageData === null) {
            return null;
        }

        $model = new PictureModel();
        $model->Name = $name;
        $model->UpperLeftColumn = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $model->LowerRightColumnOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->LowerRightRow = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $model->LowerRightRowOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->ImageData = $imageData;
        $model->ImageExtension = Picture::DetectExtension($imageData);
        $model->OriginalRId = $rId;

        self::LoadSpPrExtents($pic, $model);
        return $model;
    }

    private static function ParseOneCellPicture(XElement $pic, XElement $fromEl, $extEl, $imageTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $rId = self::GetBlipRId($pic);
        if ($rId === null || $rId === '') {
            return null;
        }

        $name = self::GetPictureName($pic);
        $imageData = self::LoadImageData($rId, $imageTargets, $archive, $diagnostics, $options, $sheetName, $name);
        if ($imageData === null) {
            return null;
        }

        $fromCol = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
        $fromRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
        $cx = 0;
        $cy = 0;

        if ($extEl !== null) {
            $cxText = (string) $extEl->Attribute('cx');
            if ($cxText === '') {
                $cxText = '0';
            }
            if (is_numeric($cxText)) {
                $cx = (int) $cxText;
            }

            $cyText = (string) $extEl->Attribute('cy');
            if ($cyText === '') {
                $cyText = '0';
            }
            if (is_numeric($cyText)) {
                $cy = (int) $cyText;
            }
        }

        $colSpan = $cx > 0 ? (int) ($cx / 609600) + 1 : 1;
        $rowSpan = $cy > 0 ? (int) ($cy / 190500) + 1 : 1;

        $model = new PictureModel();
        $model->Name = $name;
        $model->UpperLeftColumn = $fromCol;
        $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
        $model->UpperLeftRow = $fromRow;
        $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        $model->LowerRightColumn = $fromCol + $colSpan;
        $model->LowerRightRow = $fromRow + $rowSpan;
        $model->ExtentCx = $cx;
        $model->ExtentCy = $cy;
        $model->ImageData = $imageData;
        $model->ImageExtension = Picture::DetectExtension($imageData);
        $model->OriginalRId = $rId;
        return $model;
    }

    private static function GetBlipRId(XElement $pic)
    {
        $blipFill = $pic->Element(XlsxWorkbookPictures::XdrNs . 'blipFill');
        if ($blipFill === null) {
            return null;
        }

        $blip = $blipFill->Element(XlsxWorkbookPictures::ANs . 'blip');
        if ($blip === null) {
            return null;
        }

        return (string) $blip->Attribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'embed');
    }

    private static function GetPictureName(XElement $pic)
    {
        $nvPicPr = $pic->Element(XlsxWorkbookPictures::XdrNs . 'nvPicPr');
        if ($nvPicPr === null) {
            return '';
        }

        $cNvPr = $nvPicPr->Element(XlsxWorkbookPictures::XdrNs . 'cNvPr');
        return $cNvPr === null ? '' : (string) $cNvPr->Attribute('name');
    }

    private static function LoadImageData($rId, $imageTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $pictureName)
    {
        if (!array_key_exists($rId, $imageTargets)) {
            $issue = new LoadIssue('PIC-L001', DiagnosticSeverity::LossyRecoverable, "Image relationship '" . $rId . "' for picture '" . $pictureName . "' could not be resolved; picture was dropped.", true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        $mediaUri = $imageTargets[$rId];

        $mediaEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $mediaUri);
        if ($mediaEntry === null) {
            $issue = new LoadIssue('PIC-L001', DiagnosticSeverity::LossyRecoverable, "Image media file '" . $mediaUri . "' was not found in the package; picture was dropped.", true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        return $mediaEntry->getContent();
    }

    private static function LoadSpPrExtents(XElement $pic, PictureModel $model)
    {
        $spPr = $pic->Element(XlsxWorkbookPictures::XdrNs . 'spPr');
        if ($spPr === null) {
            return;
        }

        $xfrm = $spPr->Element(XlsxWorkbookPictures::ANs . 'xfrm');
        if ($xfrm === null) {
            return;
        }

        $ext = $xfrm->Element(XlsxWorkbookPictures::ANs . 'ext');
        if ($ext === null) {
            return;
        }

        $cx = 0;
        $cy = 0;
        $cxText = (string) $ext->Attribute('cx');
        if ($cxText === '') {
            $cxText = '0';
        }
        if (is_numeric($cxText)) {
            $cx = (int) $cxText;
        }

        $cyText = (string) $ext->Attribute('cy');
        if ($cyText === '') {
            $cyText = '0';
        }
        if (is_numeric($cyText)) {
            $cy = (int) $cyText;
        }

        $model->ExtentCx = $cx;
        $model->ExtentCy = $cy;
    }

    private static function LoadDrawingImageTargets(ZipArchive $archive, $drawingUri)
    {
        return self::LoadDrawingTargetsByType($archive, $drawingUri, XlsxWorkbookPictures::ImageRelationshipType);
    }

    private static function LoadDrawingChartTargets(ZipArchive $archive, $drawingUri)
    {
        $relsUri = self::GetDrawingRelsUri($drawingUri);
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $relsUri);
        if ($entry === null) {
            return array();
        }

        $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        $targets = array();
        $relationships = $document->Root !== null ? $document->Root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship') : array();
        foreach ($relationships as $rel) {
            $id = (string) $rel->Attribute('Id');
            $type = (string) $rel->Attribute('Type');
            $target = (string) $rel->Attribute('Target');
            if ($id === null || $id === '' || $target === null || $target === '') {
                continue;
            }

            $isStandard = strcasecmp($type, XlsxWorkbookPictures::ChartRelationshipType) === 0;
            $isChartEx = strcasecmp($type, XlsxWorkbookPictures::ChartExRelationshipType) === 0;
            if (!$isStandard && !$isChartEx) {
                continue;
            }

            $chartTarget = new ChartTarget();
            $chartTarget->Uri = XlsxWorkbookArchiveHelpers::ResolvePartUri($drawingUri, $target);
            $chartTarget->IsChartEx = $isChartEx;
            $targets[$id] = $chartTarget;
        }

        return $targets;
    }

    private static function LoadDrawingTargetsByType(ZipArchive $archive, $drawingUri, $relationshipType)
    {
        $relsUri = self::GetDrawingRelsUri($drawingUri);
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $relsUri);
        if ($entry === null) {
            return array();
        }

        $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        $targets = array();
        $relationships = $document->Root !== null ? $document->Root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship') : array();
        foreach ($relationships as $rel) {
            $id = (string) $rel->Attribute('Id');
            $type = (string) $rel->Attribute('Type');
            $target = (string) $rel->Attribute('Target');
            if ($id === null || $id === '' || $target === null || $target === '') {
                continue;
            }

            if (strcasecmp($type, $relationshipType) !== 0) {
                continue;
            }

            $targets[$id] = XlsxWorkbookArchiveHelpers::ResolvePartUri($drawingUri, $target);
        }

        return $targets;
    }

    private static function FindDrawingUri(ZipArchive $archive, $worksheetUri)
    {
        $relsUri = self::GetWorksheetRelsUri($worksheetUri);
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $relsUri);
        if ($entry === null) {
            return '';
        }

        $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        if ($document->Root === null) {
            return '';
        }

        foreach ($document->Root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship') as $rel) {
            $type = (string) $rel->Attribute('Type');
            $target = (string) $rel->Attribute('Target');
            if (strcasecmp($type, XlsxWorkbookPictures::DrawingRelationshipType) === 0 && !($target === null || $target === '')) {
                return XlsxWorkbookArchiveHelpers::ResolvePartUri($worksheetUri, $target);
            }
        }

        return '';
    }

    private static function ParseAnchorInt($element)
    {
        if ($element === null) {
            return 0;
        }

        $value = 0;
        if (is_numeric($element->Value)) {
            $value = (int) $element->Value;
        }
        return $value;
    }

    private static function ParseAnchorLong($element)
    {
        if ($element === null) {
            return 0;
        }

        $value = 0;
        if (is_numeric($element->Value)) {
            $value = (int) $element->Value;
        }
        return $value;
    }

    private static function GetWorksheetRelsUri($worksheetUri)
    {
        $normalized = ltrim($worksheetUri, '/');
        $slashIndex = strrpos($normalized, '/');
        $directory = $slashIndex !== false ? substr($normalized, 0, $slashIndex + 1) : '';
        $fileName = $slashIndex !== false ? substr($normalized, $slashIndex + 1) : $normalized;
        return '/' . $directory . '_rels/' . $fileName . '.rels';
    }

    private static function GetDrawingRelsUri($drawingUri)
    {
        $normalized = ltrim($drawingUri, '/');
        $slashIndex = strrpos($normalized, '/');
        $directory = $slashIndex !== false ? substr($normalized, 0, $slashIndex + 1) : '';
        $fileName = $slashIndex !== false ? substr($normalized, $slashIndex + 1) : $normalized;
        return '/' . $directory . '_rels/' . $fileName . '.rels';
    }

    // --- Chart loading ---

    const McNamespace = '{http://schemas.openxmlformats.org/markup-compatibility/2006}';
    const ChartExGraphicDataUri = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

    private static function LoadTwoCellAnchorCharts(WorksheetModel $worksheetModel, XElement $drawingRoot, $chartTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        foreach ($drawingRoot->Elements(XlsxWorkbookPictures::XdrNs . 'twoCellAnchor') as $anchor) {
            $fromEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'from');
            $toEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'to');
            if ($fromEl === null || $toEl === null) {
                continue;
            }

            // Standard chart: direct graphicFrame child
            $gf = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'graphicFrame');
            if ($gf !== null && self::IsChartGraphicFrame($gf)) {
                $model = self::ParseChart($gf, null, $fromEl, $toEl, null, $chartTargets, $archive, $diagnostics, $options, $sheetName);
                if ($model !== null) {
                    $worksheetModel->Charts[] = $model;
                }
                continue;
            }

            // ChartEx: mc:AlternateContent wrapping a graphicFrame with chartex URI
            $altContent = $anchor->Element(self::McNamespace . 'AlternateContent');
            if ($altContent !== null) {
                $choice = $altContent->Element(self::McNamespace . 'Choice');
                $choiceGf = $choice === null ? null : $choice->Element(XlsxWorkbookPictures::XdrNs . 'graphicFrame');
                if ($choiceGf !== null && self::IsChartExGraphicFrame($choiceGf)) {
                    $model = self::ParseChart($choiceGf, $altContent, $fromEl, $toEl, null, $chartTargets, $archive, $diagnostics, $options, $sheetName);
                    if ($model !== null) {
                        $worksheetModel->Charts[] = $model;
                    }
                }
            }
        }
    }

    private static function LoadOneCellAnchorCharts(WorksheetModel $worksheetModel, XElement $drawingRoot, $chartTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        foreach ($drawingRoot->Elements(XlsxWorkbookPictures::XdrNs . 'oneCellAnchor') as $anchor) {
            $fromEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'from');
            $extEl = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'ext');
            if ($fromEl === null) {
                continue;
            }

            $gf = $anchor->Element(XlsxWorkbookPictures::XdrNs . 'graphicFrame');
            if ($gf !== null && self::IsChartGraphicFrame($gf)) {
                $model = self::ParseChart($gf, null, $fromEl, null, $extEl, $chartTargets, $archive, $diagnostics, $options, $sheetName);
                if ($model !== null) {
                    $worksheetModel->Charts[] = $model;
                }
                continue;
            }

            $altContent = $anchor->Element(self::McNamespace . 'AlternateContent');
            if ($altContent !== null) {
                $choice = $altContent->Element(self::McNamespace . 'Choice');
                $choiceGf = $choice === null ? null : $choice->Element(XlsxWorkbookPictures::XdrNs . 'graphicFrame');
                if ($choiceGf !== null && self::IsChartExGraphicFrame($choiceGf)) {
                    $model = self::ParseChart($choiceGf, $altContent, $fromEl, null, $extEl, $chartTargets, $archive, $diagnostics, $options, $sheetName);
                    if ($model !== null) {
                        $worksheetModel->Charts[] = $model;
                    }
                }
            }
        }
    }

    private static function IsChartGraphicFrame(XElement $graphicFrame)
    {
        $graphic = $graphicFrame->Element(XlsxWorkbookPictures::ANs . 'graphic');
        $graphicData = $graphic === null ? null : $graphic->Element(XlsxWorkbookPictures::ANs . 'graphicData');
        if ($graphicData === null) {
            return false;
        }

        $uri = (string) $graphicData->Attribute('uri');
        return strcasecmp($uri, XlsxWorkbookPictures::ChartGraphicDataUri) === 0;
    }

    private static function IsChartExGraphicFrame(XElement $graphicFrame)
    {
        $graphic = $graphicFrame->Element(XlsxWorkbookPictures::ANs . 'graphic');
        $graphicData = $graphic === null ? null : $graphic->Element(XlsxWorkbookPictures::ANs . 'graphicData');
        if ($graphicData === null) {
            return false;
        }

        $uri = (string) $graphicData->Attribute('uri');
        return strcasecmp($uri, self::ChartExGraphicDataUri) === 0;
    }

    const ChartExNs = '{http://schemas.microsoft.com/office/drawing/2014/chartex}';

    private static function ParseChart(XElement $graphicFrame, $rawContainerElement, $fromEl, $toEl, $extEl, $chartTargets, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $graphic = $graphicFrame->Element(XlsxWorkbookPictures::ANs . 'graphic');
        $graphicData = $graphic === null ? null : $graphic->Element(XlsxWorkbookPictures::ANs . 'graphicData');

        // Standard chart: c:chart element; ChartEx: cx:chart element
        $chartEl = null;
        if ($graphicData !== null) {
            $chartEl = $graphicData->Element(XlsxWorkbookPictures::ChartNs . 'chart');
            if ($chartEl === null) {
                $chartEl = $graphicData->Element(self::ChartExNs . 'chart');
            }
        }
        $rId = $chartEl !== null ? (string) $chartEl->Attribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id') : null;
        if ($rId === null || $rId === '') {
            return null;
        }

        if (!array_key_exists($rId, $chartTargets)) {
            $issue = new LoadIssue('CHT-L001', DiagnosticSeverity::LossyRecoverable, "Chart relationship '" . $rId . "' could not be resolved; chart was dropped.", true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        $chartTarget = $chartTargets[$rId];
        $chartUri = $chartTarget->Uri;
        $isChartEx = $chartTarget->IsChartEx;

        $chartEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $chartUri);
        if ($chartEntry === null) {
            $issue = new LoadIssue('CHT-R001', DiagnosticSeverity::Recoverable, ($isChartEx ? 'ChartEx' : 'Chart') . " part '" . $chartUri . "' was not found; chart was skipped.", true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return null;
        }

        $rawChartXml = $chartEntry->getContent();

        $model = new ChartModel();
        $model->Name = self::GetChartName($graphicFrame);
        $model->RawChartXml = $rawChartXml;
        $model->ChartType = self::DetectChartType($rawChartXml);
        $model->IsChartEx = $isChartEx;
        $model->OriginalRId = $rId;
        $model->RawGraphicFrameXml = $rawContainerElement !== null ? $rawContainerElement->toString() : null;

        if ($toEl !== null) {
            $model->UpperLeftColumn = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
            $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
            $model->UpperLeftRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
            $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
            $model->LowerRightColumn = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
            $model->LowerRightColumnOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
            $model->LowerRightRow = self::ParseAnchorInt($toEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
            $model->LowerRightRowOffset = self::ParseAnchorLong($toEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
        } elseif ($extEl !== null) {
            $fromCol = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'col'));
            $fromRow = self::ParseAnchorInt($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'row'));
            $cx = 0;
            $cy = 0;
            $cxText = (string) $extEl->Attribute('cx');
            if ($cxText === '') {
                $cxText = '0';
            }
            if (is_numeric($cxText)) {
                $cx = (int) $cxText;
            }
            $cyText = (string) $extEl->Attribute('cy');
            if ($cyText === '') {
                $cyText = '0';
            }
            if (is_numeric($cyText)) {
                $cy = (int) $cyText;
            }
            $model->UpperLeftColumn = $fromCol;
            $model->UpperLeftColumnOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'colOff'));
            $model->UpperLeftRow = $fromRow;
            $model->UpperLeftRowOffset = self::ParseAnchorLong($fromEl->Element(XlsxWorkbookPictures::XdrNs . 'rowOff'));
            $model->LowerRightColumn = $fromCol + ($cx > 0 ? (int) ($cx / 609600) + 1 : 1);
            $model->LowerRightRow = $fromRow + ($cy > 0 ? (int) ($cy / 190500) + 1 : 1);
            $model->ExtentCx = $cx;
            $model->ExtentCy = $cy;
        }

        self::LoadChartCompanionFiles($model, $chartUri, $archive, $diagnostics, $options, $sheetName);
        return $model;
    }

    private static function GetChartName(XElement $graphicFrame)
    {
        $nvGraphicFramePr = $graphicFrame->Element(XlsxWorkbookPictures::XdrNs . 'nvGraphicFramePr');
        $cNvPr = $nvGraphicFramePr === null ? null : $nvGraphicFramePr->Element(XlsxWorkbookPictures::XdrNs . 'cNvPr');
        return $cNvPr === null ? '' : (string) $cNvPr->Attribute('name');
    }

    private static function DetectChartType($rawXml)
    {
        if ($rawXml === null || $rawXml === '') {
            return ChartType::Unknown;
        }

        if (strpos($rawXml, '<c:barChart') !== false) {
            return strpos($rawXml, 'barDir val="bar"') !== false
                ? ChartType::Bar
                : ChartType::Column;
        }

        if (strpos($rawXml, '<c:bar3DChart') !== false) {
            return strpos($rawXml, 'barDir val="bar"') !== false
                ? ChartType::Bar3D
                : ChartType::Column3D;
        }

        if (strpos($rawXml, '<c:lineChart') !== false) return ChartType::Line;
        if (strpos($rawXml, '<c:line3DChart') !== false) return ChartType::Line3D;
        if (strpos($rawXml, '<c:areaChart') !== false) return ChartType::Area;
        if (strpos($rawXml, '<c:area3DChart') !== false) return ChartType::Area3D;
        if (strpos($rawXml, '<c:pie3DChart') !== false) return ChartType::Pie3D;
        if (strpos($rawXml, '<c:pieChart') !== false) return ChartType::Pie;
        if (strpos($rawXml, '<c:doughnutChart') !== false) return ChartType::Doughnut;
        if (strpos($rawXml, '<c:scatterChart') !== false) return ChartType::Scatter;
        if (strpos($rawXml, '<c:bubbleChart') !== false) return ChartType::Bubble;
        if (strpos($rawXml, '<c:radarChart') !== false) return ChartType::Radar;
        if (strpos($rawXml, '<c:stockChart') !== false) return ChartType::Stock;
        if (strpos($rawXml, '<c:surface3DChart') !== false) {
            return strpos($rawXml, 'wireframe val="1"') !== false
                ? ChartType::SurfaceWireframe3D
                : ChartType::Surface3D;
        }

        if (strpos($rawXml, '<c:surfaceChart') !== false) return ChartType::Contour;

        return ChartType::Unknown;
    }

    private static function LoadChartCompanionFiles(ChartModel $model, $chartUri, ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $normalized = ltrim($chartUri, '/');
        $slashIndex = strrpos($normalized, '/');
        $directory = $slashIndex !== false ? substr($normalized, 0, $slashIndex + 1) : '';
        $fileName = $slashIndex !== false ? substr($normalized, $slashIndex + 1) : $normalized;
        $chartRelsUri = '/' . $directory . '_rels/' . $fileName . '.rels';

        $relsEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $chartRelsUri);
        if ($relsEntry === null) {
            return;
        }

        $relsDocument = XlsxWorkbookArchiveHelpers::LoadDocument($relsEntry);
        if ($relsDocument->Root === null) {
            return;
        }

        foreach ($relsDocument->Root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship') as $rel) {
            $relId = (string) $rel->Attribute('Id');
            $type = (string) $rel->Attribute('Type');
            $target = (string) $rel->Attribute('Target');
            if ($relId === null || $relId === '' || $target === null || $target === '') {
                continue;
            }

            $isXmlCompanion =
                strcasecmp($type, XlsxWorkbookPictures::ChartStyleRelationshipType) === 0 ||
                strcasecmp($type, XlsxWorkbookPictures::ChartColorStyleRelationshipType) === 0 ||
                strcasecmp($type, XlsxWorkbookPictures::ChartUserShapesRelationshipType) === 0;

            if ($isXmlCompanion) {
                $companionUri = XlsxWorkbookArchiveHelpers::ResolvePartUri($chartUri, $target);
                $companionEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $companionUri);
                if ($companionEntry === null) {
                    continue;
                }

                $rawContent = $companionEntry->getContent();

                // Store the original relative target path so cross-directory paths (e.g.
                // "../drawings/drawing4.xml") can be resolved correctly on save without renumbering.
                $companion = new ChartCompanionFile();
                $companion->RelationshipId = $relId;
                $companion->RelationshipType = $type;
                $companion->FileName = $target;
                $companion->RawContent = $rawContent;
                $model->CompanionFiles[] = $companion;
            } elseif (strcasecmp($type, XlsxWorkbookPictures::ImageRelationshipType) === 0) {
                $issue = new LoadIssue('CHT-R002', DiagnosticSeverity::Recoverable, "Chart '" . $model->Name . "' contains an embedded image companion; this is not yet supported and the image was not loaded.", false);
                $issue->SheetName = $sheetName;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            }
        }
    }
}

// Represents the resolved target of a drawing chart relationship.
final class ChartTarget
{
    public $Uri;
    public $IsChartEx;
}
