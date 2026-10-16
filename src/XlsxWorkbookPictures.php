<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ChartModel;
use Aspose\Cells_FOSS\Core\PictureModel;
use Aspose\Cells_FOSS\Core\ShapeModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

// Builds drawing parts (pictures, shapes, charts) for the XLSX workbook writer.
final class XlsxWorkbookPictures
{
    const DrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
    const ImageRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
    const DrawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml';
    const ChartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
    const ChartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
    const ChartExRelationshipType = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
    const ChartExContentType = 'application/vnd.ms-office.chartex+xml';
    const ChartStyleContentType = 'application/vnd.ms-office.chartstyle+xml';
    const ChartColorStyleContentType = 'application/vnd.ms-office.chartcolorstyle+xml';
    const ChartGraphicDataUri = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const ChartStyleRelationshipType = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
    const ChartColorStyleRelationshipType = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
    const ChartUserShapesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes';
    const ChartUserShapesContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml';

    const XdrNs = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';
    const ANs = '{http://schemas.openxmlformats.org/drawingml/2006/main}';
    const ChartNs = '{http://schemas.openxmlformats.org/drawingml/2006/chart}';

    private function __construct()
    {
    }

    public static function BuildDrawingDocument(WorksheetModel $worksheet, $pictureFileOffset, $chartFileOffset)
    {
        $root = new XElement(
            self::XdrNs . 'wsDr',
            new XAttribute(XNamespace::Xmlns . 'xdr', self::XdrNs),
            new XAttribute(XNamespace::Xmlns . 'a', self::ANs),
            new XAttribute(XNamespace::Xmlns . 'r', XlsxWorkbookSerializerCommon::RelationshipNs)
        );

        // Build a mapping from each picture's/shape-image's original rId to its new assigned rId
        // so that group shapes whose raw XML contains r:embed references stay correct.
        $imageCount = count($worksheet->Pictures) + count($worksheet->ShapeImages);
        $rIdMap = self::BuildImageRIdMap($worksheet->Pictures, $worksheet->ShapeImages);

        for ($i = 0; $i < count($worksheet->Pictures); $i++) {
            $picture = $worksheet->Pictures[$i];
            $rId = 'rId' . ($i + 1);
            $picId = $i + 1;
            $root->Add(self::BuildTwoCellAnchor($picture, $rId, $picId));
        }

        for ($j = 0; $j < count($worksheet->Shapes); $j++) {
            $shape = $worksheet->Shapes[$j];
            $shapeId = count($worksheet->Pictures) + $j + 1;
            $root->Add(self::BuildShapeAnchor($shape, $shapeId, $rIdMap));
        }

        for ($k = 0; $k < count($worksheet->Charts); $k++) {
            $chart = $worksheet->Charts[$k];
            $rId = 'rId' . ($imageCount + $k + 1);
            $chartId = $imageCount + count($worksheet->Shapes) + $k + 1;
            $root->Add(self::BuildChartAnchor($chart, $rId, $chartId));
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $root);
    }

    private static function BuildImageRIdMap($pictures, $shapeImages)
    {
        $map = array();
        for ($i = 0; $i < count($pictures); $i++) {
            $originalRId = $pictures[$i]->OriginalRId;
            if (!($originalRId === null || $originalRId === '')) {
                $newRId = 'rId' . ($i + 1);
                if ($originalRId !== $newRId) {
                    $map[$originalRId] = $newRId;
                }
            }
        }
        for ($j = 0; $j < count($shapeImages); $j++) {
            $originalRId = $shapeImages[$j]->OriginalRId;
            if (!($originalRId === null || $originalRId === '')) {
                $newRId = 'rId' . (count($pictures) + $j + 1);
                if ($originalRId !== $newRId) {
                    $map[$originalRId] = $newRId;
                }
            }
        }
        return $map;
    }

    private static function BuildTwoCellAnchor(PictureModel $picture, $rId, $picId)
    {
        $anchor = new XElement(
            self::XdrNs . 'twoCellAnchor',
            new XAttribute('editAs', 'oneCell')
        );

        $anchor->Add(self::BuildFromElement(
            $picture->UpperLeftColumn,
            $picture->UpperLeftColumnOffset,
            $picture->UpperLeftRow,
            $picture->UpperLeftRowOffset
        ));
        $anchor->Add(self::BuildToElement(
            $picture->LowerRightColumn,
            $picture->LowerRightColumnOffset,
            $picture->LowerRightRow,
            $picture->LowerRightRowOffset
        ));
        $anchor->Add(self::BuildPicElement($picture, $rId, $picId));
        $anchor->Add(new XElement(self::XdrNs . 'clientData'));
        return $anchor;
    }

    private static function BuildFromElement($col, $colOff, $row, $rowOff)
    {
        return new XElement(
            self::XdrNs . 'from',
            new XElement(self::XdrNs . 'col', (string) $col),
            new XElement(self::XdrNs . 'colOff', (string) $colOff),
            new XElement(self::XdrNs . 'row', (string) $row),
            new XElement(self::XdrNs . 'rowOff', (string) $rowOff)
        );
    }

    private static function BuildToElement($col, $colOff, $row, $rowOff)
    {
        return new XElement(
            self::XdrNs . 'to',
            new XElement(self::XdrNs . 'col', (string) $col),
            new XElement(self::XdrNs . 'colOff', (string) $colOff),
            new XElement(self::XdrNs . 'row', (string) $row),
            new XElement(self::XdrNs . 'rowOff', (string) $rowOff)
        );
    }

    private static function BuildPicElement(PictureModel $picture, $rId, $picId)
    {
        $cx = $picture->ExtentCx > 0
            ? $picture->ExtentCx
            : ($picture->LowerRightColumn - $picture->UpperLeftColumn) * 609600;
        $cy = $picture->ExtentCy > 0
            ? $picture->ExtentCy
            : ($picture->LowerRightRow - $picture->UpperLeftRow) * 190500;

        if ($cx <= 0) {
            $cx = 609600;
        }

        if ($cy <= 0) {
            $cy = 190500;
        }

        $pic = new XElement(self::XdrNs . 'pic');

        $nvPicPr = new XElement(
            self::XdrNs . 'nvPicPr',
            new XElement(
                self::XdrNs . 'cNvPr',
                new XAttribute('id', (string) $picId),
                new XAttribute('name', ($picture->Name === null || $picture->Name === '') ? 'Picture ' . $picId : $picture->Name)
            ),
            new XElement(
                self::XdrNs . 'cNvPicPr',
                new XElement(
                    self::ANs . 'picLocks',
                    new XAttribute('noChangeAspect', '1')
                )
            )
        );
        $pic->Add($nvPicPr);

        $blipFill = new XElement(
            self::XdrNs . 'blipFill',
            new XElement(
                self::ANs . 'blip',
                new XAttribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'embed', $rId)
            ),
            new XElement(
                self::ANs . 'stretch',
                new XElement(self::ANs . 'fillRect')
            )
        );
        $pic->Add($blipFill);

        $spPr = new XElement(
            self::XdrNs . 'spPr',
            new XElement(
                self::ANs . 'xfrm',
                new XElement(
                    self::ANs . 'off',
                    new XAttribute('x', '0'),
                    new XAttribute('y', '0')
                ),
                new XElement(
                    self::ANs . 'ext',
                    new XAttribute('cx', (string) $cx),
                    new XAttribute('cy', (string) $cy)
                )
            ),
            new XElement(
                self::ANs . 'prstGeom',
                new XAttribute('prst', 'rect')
            )
        );
        $pic->Add($spPr);

        return $pic;
    }

    private static function BuildShapeAnchor(ShapeModel $shape, $shapeId, $pictureRIdMap = null)
    {
        $anchor = new XElement(
            self::XdrNs . 'twoCellAnchor',
            new XAttribute('editAs', 'oneCell')
        );

        $anchor->Add(self::BuildFromElement(
            $shape->UpperLeftColumn,
            $shape->UpperLeftColumnOffset,
            $shape->UpperLeftRow,
            $shape->UpperLeftRowOffset
        ));
        $anchor->Add(self::BuildToElement(
            $shape->LowerRightColumn,
            $shape->LowerRightColumnOffset,
            $shape->LowerRightRow,
            $shape->LowerRightRowOffset
        ));

        if (!($shape->RawElementXml === null || $shape->RawElementXml === '')) {
            try {
                $rawXml = self::ApplyRIdRemap($shape->RawElementXml, $pictureRIdMap);
                $anchor->Add(XElement::Parse($rawXml));
            } catch (\Throwable $exception) {
                $anchor->Add(self::BuildSpElement($shape, $shapeId));
            }
        } else {
            $anchor->Add(self::BuildSpElement($shape, $shapeId));
        }

        $anchor->Add(new XElement(self::XdrNs . 'clientData'));
        return $anchor;
    }

    private static function ApplyRIdRemap($rawXml, $rIdMap)
    {
        if ($rIdMap === null || count($rIdMap) === 0 || $rawXml === null || $rawXml === '') {
            return $rawXml;
        }

        foreach ($rIdMap as $key => $value) {
            $rawXml = str_replace('"' . $key . '"', '"' . $value . '"', $rawXml);
        }

        return $rawXml;
    }

    private static function BuildSpElement(ShapeModel $shape, $shapeId)
    {
        $cx = $shape->ExtentCx > 0
            ? $shape->ExtentCx
            : ($shape->LowerRightColumn - $shape->UpperLeftColumn) * 609600;
        $cy = $shape->ExtentCy > 0
            ? $shape->ExtentCy
            : ($shape->LowerRightRow - $shape->UpperLeftRow) * 190500;

        if ($cx <= 0) {
            $cx = 609600;
        }

        if ($cy <= 0) {
            $cy = 190500;
        }

        $geomType = ($shape->GeometryType === null || $shape->GeometryType === '') ? 'rect' : $shape->GeometryType;
        $displayName = ($shape->Name === null || $shape->Name === '') ? 'Shape ' . $shapeId : $shape->Name;

        $sp = new XElement(
            self::XdrNs . 'sp',
            new XAttribute('macro', ''),
            new XAttribute('textlink', '')
        );

        $sp->Add(new XElement(
            self::XdrNs . 'nvSpPr',
            new XElement(
                self::XdrNs . 'cNvPr',
                new XAttribute('id', (string) $shapeId),
                new XAttribute('name', $displayName)
            ),
            new XElement(self::XdrNs . 'cNvSpPr')
        ));

        $sp->Add(new XElement(
            self::XdrNs . 'spPr',
            new XElement(
                self::ANs . 'xfrm',
                new XElement(
                    self::ANs . 'off',
                    new XAttribute('x', '0'),
                    new XAttribute('y', '0')
                ),
                new XElement(
                    self::ANs . 'ext',
                    new XAttribute('cx', (string) $cx),
                    new XAttribute('cy', (string) $cy)
                )
            ),
            new XElement(
                self::ANs . 'prstGeom',
                new XAttribute('prst', $geomType),
                new XElement(self::ANs . 'avLst')
            )
        ));

        if (!($shape->RawStyleXml === null || $shape->RawStyleXml === '')) {
            try {
                $sp->Add(XElement::Parse($shape->RawStyleXml));
            } catch (\Throwable $exception) {
            }
        } else {
            $sp->Add(self::BuildDefaultShapeStyle());
        }

        if (!($shape->RawTxBodyXml === null || $shape->RawTxBodyXml === '')) {
            try {
                $sp->Add(XElement::Parse($shape->RawTxBodyXml));
            } catch (\Throwable $exception) {
                $sp->Add(self::BuildMinimalTxBody());
            }
        } else {
            $sp->Add(self::BuildMinimalTxBody());
        }

        return $sp;
    }

    private static function BuildMinimalTxBody()
    {
        return new XElement(
            self::XdrNs . 'txBody',
            new XElement(self::ANs . 'bodyPr'),
            new XElement(self::ANs . 'lstStyle'),
            new XElement(self::ANs . 'p')
        );
    }

    private static function BuildDefaultShapeStyle()
    {
        return new XElement(
            self::XdrNs . 'style',
            new XElement(self::ANs . 'lnRef', new XAttribute('idx', '2'), new XElement(self::ANs . 'schemeClr', new XAttribute('val', 'accent1'), new XElement(self::ANs . 'shade', new XAttribute('val', '50000')))),
            new XElement(self::ANs . 'fillRef', new XAttribute('idx', '1'), new XElement(self::ANs . 'schemeClr', new XAttribute('val', 'accent1'))),
            new XElement(self::ANs . 'effectRef', new XAttribute('idx', '0'), new XElement(self::ANs . 'schemeClr', new XAttribute('val', 'accent1'))),
            new XElement(self::ANs . 'fontRef', new XAttribute('idx', 'minor'), new XElement(self::ANs . 'schemeClr', new XAttribute('val', 'lt1')))
        );
    }

    public static function BuildDrawingRelationshipsDocument(WorksheetModel $worksheet, $pictureFileOffset, $chartFileOffset)
    {
        $relationships = new XElement(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationships');
        $imageCount = count($worksheet->Pictures) + count($worksheet->ShapeImages);

        for ($i = 0; $i < count($worksheet->Pictures); $i++) {
            $picture = $worksheet->Pictures[$i];
            $globalPictureNumber = $pictureFileOffset + $i + 1;
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . ($i + 1)),
                new XAttribute('Type', self::ImageRelationshipType),
                new XAttribute('Target', '../media/image' . $globalPictureNumber . '.' . $picture->ImageExtension)
            ));
        }

        // Shape-embedded images follow top-level pictures so their rIds come after pictures
        // but before charts. This keeps r:embed references in group shape raw XML valid.
        for ($j = 0; $j < count($worksheet->ShapeImages); $j++) {
            $shapeImage = $worksheet->ShapeImages[$j];
            $globalPictureNumber = $pictureFileOffset + count($worksheet->Pictures) + $j + 1;
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . (count($worksheet->Pictures) + $j + 1)),
                new XAttribute('Type', self::ImageRelationshipType),
                new XAttribute('Target', '../media/image' . $globalPictureNumber . '.' . $shapeImage->Extension)
            ));
        }

        for ($k = 0; $k < count($worksheet->Charts); $k++) {
            $chart = $worksheet->Charts[$k];
            $globalChartNumber = $chartFileOffset + $k + 1;
            $relType = $chart->IsChartEx ? self::ChartExRelationshipType : self::ChartRelationshipType;
            $fileName = $chart->IsChartEx
                ? 'chartEx' . $globalChartNumber . '.xml'
                : 'chart' . $globalChartNumber . '.xml';
            $relationships->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . ($imageCount + $k + 1)),
                new XAttribute('Type', $relType),
                new XAttribute('Target', '../charts/' . $fileName)
            ));
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $relationships);
    }

    public static function WritePictureMediaEntries(ZipArchive $archive, WorksheetModel $worksheet, $pictureFileOffset)
    {
        for ($i = 0; $i < count($worksheet->Pictures); $i++) {
            $picture = $worksheet->Pictures[$i];
            $globalPictureNumber = $pictureFileOffset + $i + 1;
            $path = 'xl/media/image' . $globalPictureNumber . '.' . $picture->ImageExtension;
            $archive->addEntry($path, $picture->ImageData);
        }

        for ($j = 0; $j < count($worksheet->ShapeImages); $j++) {
            $shapeImage = $worksheet->ShapeImages[$j];
            $globalPictureNumber = $pictureFileOffset + count($worksheet->Pictures) + $j + 1;
            $path = 'xl/media/image' . $globalPictureNumber . '.' . $shapeImage->Extension;
            $archive->addEntry($path, $shapeImage->ImageData);
        }
    }

    private static function BuildChartAnchor(ChartModel $chart, $rId, $chartId)
    {
        $anchor = new XElement(
            self::XdrNs . 'twoCellAnchor',
            new XAttribute('editAs', 'oneCell')
        );

        $anchor->Add(self::BuildFromElement(
            $chart->UpperLeftColumn,
            $chart->UpperLeftColumnOffset,
            $chart->UpperLeftRow,
            $chart->UpperLeftRowOffset
        ));
        $anchor->Add(self::BuildToElement(
            $chart->LowerRightColumn,
            $chart->LowerRightColumnOffset,
            $chart->LowerRightRow,
            $chart->LowerRightRowOffset
        ));

        if ($chart->IsChartEx && !($chart->RawGraphicFrameXml === null || $chart->RawGraphicFrameXml === '')) {
            // Substitute the new rId into the preserved raw element (mc:AlternateContent or graphicFrame)
            $updatedXml = ($chart->OriginalRId === null || $chart->OriginalRId === '')
                ? $chart->RawGraphicFrameXml
                : str_replace('"' . $chart->OriginalRId . '"', '"' . $rId . '"', $chart->RawGraphicFrameXml);
            try {
                $anchor->Add(XElement::Parse($updatedXml));
            } catch (\Throwable $exception) {
                // If raw XML is unparseable preserve a minimal placeholder so the file remains valid
            }
        } else {
            $anchor->Add(self::BuildGraphicFrameElement($chart, $rId, $chartId));
        }

        $anchor->Add(new XElement(self::XdrNs . 'clientData'));
        return $anchor;
    }

    private static function BuildGraphicFrameElement(ChartModel $chart, $rId, $frameId)
    {
        $displayName = ($chart->Name === null || $chart->Name === '')
            ? 'Chart ' . $frameId
            : $chart->Name;

        $graphicFrame = new XElement(
            self::XdrNs . 'graphicFrame',
            new XAttribute('macro', '')
        );

        $graphicFrame->Add(new XElement(
            self::XdrNs . 'nvGraphicFramePr',
            new XElement(
                self::XdrNs . 'cNvPr',
                new XAttribute('id', (string) $frameId),
                new XAttribute('name', $displayName)
            ),
            new XElement(self::XdrNs . 'cNvGraphicFramePr')
        ));

        $graphicFrame->Add(new XElement(
            self::XdrNs . 'xfrm',
            new XElement(
                self::ANs . 'off',
                new XAttribute('x', '0'),
                new XAttribute('y', '0')
            ),
            new XElement(
                self::ANs . 'ext',
                new XAttribute('cx', '0'),
                new XAttribute('cy', '0')
            )
        ));

        $graphicFrame->Add(new XElement(
            self::ANs . 'graphic',
            new XElement(
                self::ANs . 'graphicData',
                new XAttribute('uri', self::ChartGraphicDataUri),
                new XElement(
                    self::ChartNs . 'chart',
                    new XAttribute(XNamespace::Xmlns . 'c', self::ChartNs),
                    new XAttribute(XNamespace::Xmlns . 'r', XlsxWorkbookSerializerCommon::RelationshipNs),
                    new XAttribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id', $rId)
                )
            )
        ));

        return $graphicFrame;
    }
}
