<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\ChartModel;
use Aspose\Cells_FOSS\Core\ExternalLinkModel;
use Aspose\Cells_FOSS\Core\SharedStringRepository;
use Aspose\Cells_FOSS\Core\SheetVisibility;
use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookSerializer
{
    private function __construct()
    {
    }

    public static function Save(WorkbookModel $model, $stream, SaveOptions $options)
    {
        if ($options->SaveFormat !== SaveFormat::Xlsx) {
            throw new UnsupportedFeatureException("Only XLSX save is supported.");
        }

        if ($stream === null) {
            throw new WorkbookSaveException("The output stream must be writable.");
        }

        if (count($model->Worksheets) === 0) {
            throw new WorkbookSaveException("A workbook must contain at least one worksheet.");
        }

        $sharedStrings = new SharedStringRepository();
        $stylesheet = XlsxWorkbookStyles::BuildStylesheet($model);
        $corePropertiesDocument = XlsxDocumentProperties::BuildCorePropertiesDocument($model);
        $extendedPropertiesDocument = XlsxDocumentProperties::BuildExtendedPropertiesDocument($model);

        foreach ($model->Worksheets as $worksheet) {
            foreach ($worksheet->Cells as $cellKey => $cellValue) {
                if (!XlsxWorkbookSerializerCommon::ShouldPersistCell($model->DefaultStyle, $cellValue)) {
                    continue;
                }

                if ($cellValue->Kind === CellValueKind::String && $options->UseSharedStrings) {
                    $text = $cellValue->Value;
                    if ($text !== null) {
                        $sharedStrings->Intern($text);
                    }
                }
            }
        }

        $tableFileOffsets = self::ComputeTableFileOffsets($model);
        $totalTableCount = self::ComputeTotalTableCount($model);
        $pictureFileOffsets = self::ComputePictureFileOffsets($model);
        $chartFileOffsets = self::ComputeChartFileOffsets($model);
        $drawingNumbers = self::ComputeDrawingNumbers($model);
        $totalDrawingCount = self::ComputeTotalDrawingCount($model);
        $imageExtensions = self::CollectImageExtensions($model);
        $chartPartNames = self::CollectChartPartNames($model, $chartFileOffsets);
        $chartContentTypes = self::CollectChartContentTypes($model);
        $chartCompanionPartNames = self::CollectChartCompanionPartNames($model, $chartFileOffsets, $totalDrawingCount);
        $chartCompanionContentTypes = self::CollectChartCompanionContentTypes($model);
        $commentFileNumbers = self::ComputeCommentFileNumbers($model);
        $totalCommentCount = self::ComputeTotalCommentCount($model);

        $userShapesDrawingCounter = $totalDrawingCount;

        $archive = new ZipArchive($stream, 'w');
        try {
            $hasSharedStrings = count($sharedStrings->Values) > 0;
            $hasTheme = !($model->RawThemeXml === null || $model->RawThemeXml === '');
            $externalLinkBaseRId = count($model->Worksheets)
                + ($hasSharedStrings ? 1 : 0)
                + 1
                + ($hasTheme ? 1 : 0)
                + 1;
            XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, '[Content_Types].xml', XlsxWorkbookSerializerCommon::BuildContentTypes($model, $hasSharedStrings, true, $corePropertiesDocument !== null, $extendedPropertiesDocument !== null, $totalTableCount, $totalDrawingCount, $imageExtensions, $chartPartNames, $chartContentTypes, $chartCompanionPartNames, $chartCompanionContentTypes, $hasTheme, $totalCommentCount));
            XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, '_rels/.rels', XlsxWorkbookSerializerCommon::BuildRootRelationships($corePropertiesDocument !== null, $extendedPropertiesDocument !== null));
            XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/workbook.xml', XlsxWorkbookSerializerCommon::BuildWorkbook($model, $externalLinkBaseRId));
            XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/_rels/workbook.xml.rels', XlsxWorkbookSerializerCommon::BuildWorkbookRelationships($model, $hasSharedStrings, true));

            if ($hasTheme) {
                $archive->addEntry('xl/theme/theme1.xml', $model->RawThemeXml);
            }

            for ($e = 0; $e < count($model->ExternalLinks); $e++) {
                $extLink = $model->ExternalLinks[$e];
                $extN = (string)($e + 1);
                $extPath = 'xl/externalLinks/externalLink' . $extN . '.xml';
                $archive->addEntry($extPath, $extLink->RawXml);

                if (!($extLink->RawRelsXml === null || $extLink->RawRelsXml === '')) {
                    $extRelsPath = 'xl/externalLinks/_rels/externalLink' . $extN . '.xml.rels';
                    $archive->addEntry($extRelsPath, $extLink->RawRelsXml);
                }
            }

            for ($i = 0; $i < count($model->Worksheets); $i++) {
                $worksheet = $model->Worksheets[$i];
                $tableFileOffset = $tableFileOffsets[$i];
                $pictureFileOffset = $pictureFileOffsets[$i];
                $chartFileOffset = $chartFileOffsets[$i];
                $drawingNumber = $drawingNumbers[$i];
                $commentFileNumber = $commentFileNumbers[$i];
                $hasDrawing = count($worksheet->Pictures) > 0 || count($worksheet->Shapes) > 0 || count($worksheet->Charts) > 0;
                $hasComments = count($worksheet->Comments) > 0;
                $externalHyperlinkCount = self::CountExternalHyperlinks($worksheet);

                for ($t = 0; $t < count($worksheet->ListObjects); $t++) {
                    $globalTableNumber = $tableFileOffset + $t + 1;
                    $tableDocument = XlsxWorkbookTables::BuildTableDocument($worksheet->ListObjects[$t], $globalTableNumber);
                    XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/tables/table' . $globalTableNumber . '.xml', $tableDocument);
                }

                if ($hasDrawing) {
                    XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/drawings/drawing' . $drawingNumber . '.xml', XlsxWorkbookPictures::BuildDrawingDocument($worksheet, $pictureFileOffset, $chartFileOffset));
                    if (count($worksheet->Pictures) > 0 || count($worksheet->ShapeImages) > 0) {
                        XlsxWorkbookPictures::WritePictureMediaEntries($archive, $worksheet, $pictureFileOffset);
                    }

                    if (count($worksheet->Pictures) > 0 || count($worksheet->ShapeImages) > 0 || count($worksheet->Charts) > 0) {
                        XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/drawings/_rels/drawing' . $drawingNumber . '.xml.rels', XlsxWorkbookPictures::BuildDrawingRelationshipsDocument($worksheet, $pictureFileOffset, $chartFileOffset));
                    }

                    for ($k = 0; $k < count($worksheet->Charts); $k++) {
                        $globalChartNumber = $chartFileOffset + $k + 1;
                        self::WriteChartFiles($archive, $worksheet->Charts[$k], $globalChartNumber, $userShapesDrawingCounter);
                    }
                }

                if ($hasComments) {
                    XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/comments' . $commentFileNumber . '.xml', XlsxWorkbookComments::BuildCommentsDocument($worksheet));
                    XlsxWorkbookComments::WriteVmlDrawing($archive, $worksheet, $commentFileNumber);
                }

                XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/worksheets/sheet' . ($i + 1) . '.xml', XlsxWorkbookWorksheetWriter::BuildWorksheet($worksheet, $model->DefaultStyle, $model->Settings->DateSystem, $sharedStrings, $options, $stylesheet, $externalHyperlinkCount, $hasDrawing, $hasComments));

                $worksheetRelationships = XlsxWorkbookTables::BuildWorksheetRelationshipsDocument($worksheet, $tableFileOffset, $drawingNumber, $commentFileNumber);
                if ($worksheetRelationships !== null) {
                    XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/worksheets/_rels/sheet' . ($i + 1) . '.xml.rels', $worksheetRelationships);
                }
            }

            if (count($sharedStrings->Values) > 0) {
                XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/sharedStrings.xml', XlsxWorkbookSerializerCommon::BuildSharedStrings($sharedStrings));
            }

            XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'xl/styles.xml', $stylesheet->Document);

            if ($corePropertiesDocument !== null) {
                XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'docProps/core.xml', $corePropertiesDocument);
            }

            if ($extendedPropertiesDocument !== null) {
                XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, 'docProps/app.xml', $extendedPropertiesDocument);
            }
        } finally {
            $archive->close();
        }
    }

    private static function ComputeTableFileOffsets(WorkbookModel $model)
    {
        $offsets = array();
        $running = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $offsets[$i] = $running;
            $running += count($model->Worksheets[$i]->ListObjects);
        }

        return $offsets;
    }

    private static function ComputeTotalTableCount(WorkbookModel $model)
    {
        $total = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $total += count($model->Worksheets[$i]->ListObjects);
        }

        return $total;
    }

    private static function CountExternalHyperlinks(WorksheetModel $worksheet)
    {
        $count = 0;
        $ordered = XlsxWorkbookHyperlinks::GetOrderedHyperlinks($worksheet->Hyperlinks);
        for ($i = 0; $i < count($ordered); $i++) {
            if (!($ordered[$i]->Address === null || $ordered[$i]->Address === '')) {
                $count++;
            }
        }

        return $count;
    }

    private static function ComputePictureFileOffsets(WorkbookModel $model)
    {
        $offsets = array();
        $running = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $offsets[$i] = $running;
            $running += count($model->Worksheets[$i]->Pictures) + count($model->Worksheets[$i]->ShapeImages);
        }

        return $offsets;
    }

    private static function ComputeChartFileOffsets(WorkbookModel $model)
    {
        $offsets = array();
        $running = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $offsets[$i] = $running;
            $running += count($model->Worksheets[$i]->Charts);
        }

        return $offsets;
    }

    private static function ComputeDrawingNumbers(WorkbookModel $model)
    {
        $numbers = array();
        $drawingCounter = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $ws = $model->Worksheets[$i];
            if (count($ws->Pictures) > 0 || count($ws->Shapes) > 0 || count($ws->Charts) > 0) {
                $drawingCounter++;
                $numbers[$i] = $drawingCounter;
            } else {
                $numbers[$i] = 0;
            }
        }

        return $numbers;
    }

    private static function ComputeCommentFileNumbers(WorkbookModel $model)
    {
        $numbers = array();
        $counter = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            if (count($model->Worksheets[$i]->Comments) > 0) {
                $counter++;
                $numbers[$i] = $counter;
            } else {
                $numbers[$i] = 0;
            }
        }

        return $numbers;
    }

    private static function ComputeTotalCommentCount(WorkbookModel $model)
    {
        $total = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            if (count($model->Worksheets[$i]->Comments) > 0) {
                $total++;
            }
        }

        return $total;
    }

    private static function ComputeTotalDrawingCount(WorkbookModel $model)
    {
        $total = 0;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $ws = $model->Worksheets[$i];
            if (count($ws->Pictures) > 0 || count($ws->Shapes) > 0 || count($ws->Charts) > 0) {
                $total++;
            }
        }

        return $total;
    }

    private static function CollectChartCompanionPartNames(WorkbookModel $model, $chartFileOffsets, $totalDrawingCount)
    {
        $result = array();
        $userShapesCounter = $totalDrawingCount;
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $charts = $model->Worksheets[$i]->Charts;
            for ($k = 0; $k < count($charts); $k++) {
                $globalChartNumber = $chartFileOffsets[$i] + $k + 1;
                foreach ($charts[$k]->CompanionFiles as $companion) {
                    if (strpos($companion->FileName, '/') !== false) {
                        $userShapesCounter++;
                        $partName = '/xl/drawings/drawing' . (string)$userShapesCounter . '.xml';
                    } else {
                        $partName = '/' . self::ResolveCompanionPath(self::RenumberCompanionFileName($companion->FileName, $globalChartNumber));
                    }

                    $result[] = $partName;
                }
            }
        }

        return $result;
    }

    private static function CollectChartCompanionContentTypes(WorkbookModel $model)
    {
        $result = array();
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $charts = $model->Worksheets[$i]->Charts;
            for ($k = 0; $k < count($charts); $k++) {
                foreach ($charts[$k]->CompanionFiles as $companion) {
                    $result[] = self::GetCompanionContentType($companion->RelationshipType);
                }
            }
        }

        return $result;
    }

    private static function CollectChartPartNames(WorkbookModel $model, $chartFileOffsets)
    {
        $result = array();
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $charts = $model->Worksheets[$i]->Charts;
            for ($k = 0; $k < count($charts); $k++) {
                $globalChartNumber = (string)($chartFileOffsets[$i] + $k + 1);
                $fileName = $charts[$k]->IsChartEx
                    ? '/xl/charts/chartEx' . $globalChartNumber . '.xml'
                    : '/xl/charts/chart' . $globalChartNumber . '.xml';
                $result[] = $fileName;
            }
        }

        return $result;
    }

    private static function CollectChartContentTypes(WorkbookModel $model)
    {
        $result = array();
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $charts = $model->Worksheets[$i]->Charts;
            for ($k = 0; $k < count($charts); $k++) {
                $result[] = $charts[$k]->IsChartEx ? XlsxWorkbookPictures::ChartExContentType : XlsxWorkbookPictures::ChartContentType;
            }
        }

        return $result;
    }

    private static function GetCompanionContentType($relationshipType)
    {
        if (strcasecmp((string)$relationshipType, XlsxWorkbookPictures::ChartStyleRelationshipType) === 0) {
            return XlsxWorkbookPictures::ChartStyleContentType;
        }

        if (strcasecmp((string)$relationshipType, XlsxWorkbookPictures::ChartColorStyleRelationshipType) === 0) {
            return XlsxWorkbookPictures::ChartColorStyleContentType;
        }

        if (strcasecmp((string)$relationshipType, XlsxWorkbookPictures::ChartUserShapesRelationshipType) === 0) {
            return XlsxWorkbookPictures::ChartUserShapesContentType;
        }

        return "application/octet-stream";
    }

    private static function RenumberCompanionFileName($originalName, $globalNumber)
    {
        // Cross-directory paths (e.g. "../drawings/drawing4.xml") are kept verbatim
        // to avoid overwriting worksheet drawings with chart user-shape files.
        if (strpos($originalName, '/') !== false) {
            return $originalName;
        }

        $suffixStart = strrpos($originalName, '.');
        if ($suffixStart === false || $suffixStart < 0) {
            return $originalName;
        }

        $suffix = substr($originalName, $suffixStart);
        $basePart = substr($originalName, 0, $suffixStart);
        $trimmed = rtrim($basePart, '0123456789');
        if ($trimmed === null || $trimmed === '' || $trimmed === $basePart) {
            return $originalName;
        }

        return $trimmed . (string)$globalNumber . $suffix;
    }

    // Resolves a chart companion target relative to "xl/charts/" and normalises ".." segments.
    // e.g. ("../drawings/drawing4.xml") -> "xl/drawings/drawing4.xml"
    //      ("style1.xml")             -> "xl/charts/style1.xml"
    private static function ResolveCompanionPath($renumberedTarget)
    {
        $parts = explode('/', 'xl/charts/' . $renumberedTarget);
        $stack = array();
        foreach ($parts as $part) {
            if ($part === '..') {
                if (count($stack) > 0) {
                    array_splice($stack, count($stack) - 1, 1);
                }
            } elseif (strlen($part) > 0 && $part !== '.') {
                $stack[] = $part;
            }
        }

        return implode('/', $stack);
    }

    private static function WriteChartFiles(ZipArchive $archive, ChartModel $chart, $globalChartNumber, &$userShapesDrawingCounter)
    {
        $chartFileName = $chart->IsChartEx
            ? 'chartEx' . (string)$globalChartNumber . '.xml'
            : 'chart' . (string)$globalChartNumber . '.xml';
        $chartPath = 'xl/charts/' . $chartFileName;
        $archive->addEntry($chartPath, $chart->RawChartXml);

        if (count($chart->CompanionFiles) === 0) {
            return;
        }

        $companionRels = new XElement(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationships');
        foreach ($chart->CompanionFiles as $companion) {
            if (strpos($companion->FileName, '/') !== false) {
                // Cross-directory companion (e.g. chart user shapes): assign the next
                // drawing number after all worksheet drawings to avoid collisions.
                $userShapesDrawingCounter++;
                $drawingName = 'drawing' . (string)$userShapesDrawingCounter . '.xml';
                $companionPath = 'xl/drawings/' . $drawingName;
                $relsTarget = '../drawings/' . $drawingName;
            } else {
                $newFileName = self::RenumberCompanionFileName($companion->FileName, $globalChartNumber);
                $companionPath = self::ResolveCompanionPath($newFileName);
                $relsTarget = $newFileName;
            }

            $archive->addEntry($companionPath, $companion->RawContent);

            $companionRels->Add(new XElement(
                XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', $companion->RelationshipId),
                new XAttribute('Type', $companion->RelationshipType),
                new XAttribute('Target', $relsTarget)
            ));
        }

        $relsPath = 'xl/charts/_rels/' . $chartFileName . '.rels';
        XlsxWorkbookSerializerCommon::WriteXmlEntry($archive, $relsPath, new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $companionRels));
    }

    private static function CollectImageExtensions(WorkbookModel $model)
    {
        $seen = array();
        $result = array();
        for ($i = 0; $i < count($model->Worksheets); $i++) {
            $ws = $model->Worksheets[$i];
            for ($p = 0; $p < count($ws->Pictures); $p++) {
                $ext = $ws->Pictures[$p]->ImageExtension;
                $seenKey = $ext === null ? '' : strtolower($ext);
                if (!($ext === null || $ext === '') && !isset($seen[$seenKey])) {
                    $seen[$seenKey] = true;
                    $result[] = $ext;
                }
            }
            for ($s = 0; $s < count($ws->ShapeImages); $s++) {
                $ext = $ws->ShapeImages[$s]->Extension;
                $seenKey = $ext === null ? '' : strtolower($ext);
                if (!($ext === null || $ext === '') && !isset($seen[$seenKey])) {
                    $seen[$seenKey] = true;
                    $result[] = $ext;
                }
            }
        }

        return $result;
    }

    public static function Load($stream, LoadOptions $options, LoadDiagnostics $diagnostics)
    {
        if ($stream === null) {
            throw new WorkbookLoadException("The input stream must be readable.");
        }

        $workingStream = XlsxWorkbookArchiveHelpers::EnsureSeekable($stream);

        try {
            $archive = new ZipArchive($workingStream, 'r');
            try {
                $workbookEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, '/xl/workbook.xml');
                if ($workbookEntry === null) {
                    throw new InvalidFileFormatException("The package does not contain /xl/workbook.xml.");
                }

                $workbookDocument = XlsxWorkbookArchiveHelpers::LoadDocument($workbookEntry);
                $workbookRelationships = XlsxWorkbookArchiveHelpers::LoadRelationships($archive, '/xl/_rels/workbook.xml.rels', '/xl/workbook.xml');
                $sharedStrings = XlsxWorkbookArchiveHelpers::LoadSharedStrings($archive, $workbookRelationships, $options, $diagnostics);
                $stylesheet = XlsxWorkbookStyles::LoadStylesheet($archive, $workbookRelationships, $options, $diagnostics);

                $workbookModel = new WorkbookModel();
                $workbookModel->Worksheets = array();
                $workbookModel->DefaultStyle = $stylesheet->DefaultCellStyle->Clone();
                $workbookModel->RawTableStylesXml = $stylesheet->RawTableStylesXml;
                $workbookModel->TableDifferentialStyles = array_merge($workbookModel->TableDifferentialStyles, $stylesheet->DifferentialFormats);
                $workbookModel->RawTableDifferentialStylesXml = array_merge($workbookModel->RawTableDifferentialStylesXml, $stylesheet->RawDifferentialFormatsXml);

                $workbookRoot = $workbookDocument->Root;
                if ($workbookRoot === null) {
                    throw new InvalidFileFormatException("The workbook XML is empty.");
                }

                $worksheetDefinedNames = XlsxWorkbookPageSetup::LoadWorksheetDefinedNames($workbookRoot, $diagnostics, $options);

                $sheets = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'sheets');
                if ($sheets === null) {
                    throw new InvalidFileFormatException("The workbook XML does not contain a sheets element.");
                }

                $sheetElements = $sheets->Elements(XlsxWorkbookSerializerCommon::MainNs . 'sheet');
                if (count($sheetElements) === 0) {
                    throw new InvalidFileFormatException("The workbook XML does not contain any worksheets.");
                }

                XlsxWorkbookProperties::LoadWorkbookMetadata($workbookRoot, $workbookModel, count($sheetElements), $diagnostics, $options);
                XlsxDocumentProperties::LoadDocumentProperties($archive, $workbookModel, $diagnostics, $options);
                XlsxWorkbookDefinedNames::LoadWorkbookDefinedNames($workbookRoot, $workbookModel, count($sheetElements), $diagnostics, $options);
                $workbookModel->RawThemeXml = self::LoadRawTheme($archive);
                self::LoadExternalLinks($archive, $workbookRoot, $workbookRelationships, $workbookModel);

                $workbookModel->ActiveSheetIndex = $workbookModel->ActiveSheetIndex < count($sheetElements) ? $workbookModel->ActiveSheetIndex : 0;
                for ($index = 0; $index < count($sheetElements); $index++) {
                    $sheetElement = $sheetElements[$index];
                    $sheetName = (string)$sheetElement->Attribute('name');
                    if ($sheetName === null || trim($sheetName) === '') {
                        $sheetName = 'Sheet' . ($index + 1);
                    }

                    $resolvedSheetName = $sheetName;

                    $relationshipId = (string)$sheetElement->Attribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id');
                    $worksheetUri = XlsxWorkbookArchiveHelpers::ResolveWorksheetUri($index, $relationshipId, $workbookRelationships, $archive, $diagnostics, $resolvedSheetName, $options);
                    $worksheetEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $worksheetUri);
                    if ($worksheetEntry === null) {
                        throw new InvalidFileFormatException("Worksheet part '" . $worksheetUri . "' was not found.");
                    }

                    $worksheetHyperlinkTargets = XlsxWorkbookHyperlinks::LoadWorksheetHyperlinkTargets($archive, $worksheetUri);

                    $definedNamesState = null;
                    if (array_key_exists($index, $worksheetDefinedNames)) {
                        $definedNamesState = $worksheetDefinedNames[$index];
                    }
                    $worksheetModel = XlsxWorkbookWorksheetLoader::LoadWorksheet(
                        $resolvedSheetName,
                        XlsxWorkbookArchiveHelpers::LoadDocument($worksheetEntry),
                        $worksheetHyperlinkTargets,
                        $workbookModel->Settings->DateSystem,
                        $sharedStrings,
                        $stylesheet,
                        $diagnostics,
                        $options,
                        $definedNamesState,
                        $archive,
                        $worksheetUri
                    );

                    $state = (string)$sheetElement->Attribute('state');
                    switch ($state) {
                        case 'hidden':
                            $worksheetModel->Visibility = SheetVisibility::Hidden;
                            break;
                        case 'veryHidden':
                            $worksheetModel->Visibility = SheetVisibility::VeryHidden;
                            break;
                        default:
                            $worksheetModel->Visibility = SheetVisibility::Visible;
                            break;
                    }

                    $workbookModel->Worksheets[] = $worksheetModel;
                }

                return $workbookModel;
            } finally {
                $archive->close();
            }
        } catch (CellsException $exception) {
            throw $exception;
        } catch (\RuntimeException $exception) {
            throw new InvalidFileFormatException("The workbook is not a valid XLSX zip package.", $exception);
        } finally {
            if (!($workingStream === $stream)) {
                fclose($workingStream);
            }
        }
    }

    private static function LoadRawTheme(ZipArchive $archive)
    {
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, 'xl/theme/theme1.xml');
        if ($entry === null) {
            return null;
        }

        return $entry->getContent();
    }

    private static function LoadExternalLinks(ZipArchive $archive, XElement $workbookRoot, $workbookRelationships, WorkbookModel $workbookModel)
    {
        $externalReferencesElement = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'externalReferences');
        if ($externalReferencesElement === null) {
            return;
        }

        foreach ($externalReferencesElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'externalReference') as $extRef) {
            $rId = (string)$extRef->Attribute(XlsxWorkbookSerializerCommon::RelationshipNs . 'id');
            if ($rId === null || $rId === '') {
                continue;
            }

            $partUri = null;
            if (!array_key_exists($rId, $workbookRelationships)) {
                continue;
            }
            $partUri = $workbookRelationships[$rId];

            $partEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $partUri);
            if ($partEntry === null) {
                continue;
            }

            $rawXml = $partEntry->getContent();

            $normalizedUri = ltrim($partUri, '/');
            $lastSlash = strrpos($normalizedUri, '/');
            $directory = $lastSlash !== false ? substr($normalizedUri, 0, $lastSlash + 1) : '';
            $fileName = $lastSlash !== false ? substr($normalizedUri, $lastSlash + 1) : $normalizedUri;
            $relsUri = '/' . $directory . '_rels/' . $fileName . '.rels';

            $rawRelsXml = null;
            $relsEntry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $relsUri);
            if ($relsEntry !== null) {
                $rawRelsXml = $relsEntry->getContent();
            }

            $externalLink = new ExternalLinkModel();
            $externalLink->RawXml = $rawXml;
            $externalLink->RawRelsXml = $rawRelsXml;
            $workbookModel->ExternalLinks[] = $externalLink;
        }
    }
}
