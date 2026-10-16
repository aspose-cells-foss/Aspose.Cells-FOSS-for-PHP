<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellRecord;
use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\SharedStringRepository;
use Aspose\Cells_FOSS\Core\SheetVisibility;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

/**
 * Provides shared serialization helpers for the XLSX workbook writer.
 */
final class XlsxWorkbookSerializerCommon
{
    const WorksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    const SharedStringsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
    const StylesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
    const ExternalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink';
    const ExternalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml';
    const CorePropertiesRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
    const ExtendedPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';

    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';
    const PackageRelationshipNs = '{http://schemas.openxmlformats.org/package/2006/relationships}';
    const ContentTypeNs = '{http://schemas.openxmlformats.org/package/2006/content-types}';
    const XmlNs = XNamespace::Xml;

    /**
     * Gets the set of built-in number format identifiers that represent dates.
     *
     * @return array<int, bool>
     */
    public static function builtInDateFormats()
    {
        return array(
            14 => true,
            15 => true,
            16 => true,
            17 => true,
            18 => true,
            19 => true,
            20 => true,
            21 => true,
            22 => true,
            45 => true,
            46 => true,
            47 => true,
        );
    }

    /**
     * Determines whether a cell record must be persisted.
     *
     * @param StyleValue $workbookDefaultStyle The workbook default style.
     * @param CellRecord $record The cell record.
     * @return bool True when the cell must be persisted.
     */
    public static function ShouldPersistCell(StyleValue $workbookDefaultStyle, CellRecord $record)
    {
        return $record->IsExplicitlyStored
            || !($record->Formula === null || $record->Formula === '')
            || $record->Value !== null
            || $record->Kind !== CellValueKind::Blank
            || !XlsxWorkbookStyles::StylesEqual($record->Style, $workbookDefaultStyle);
    }

    /**
     * Writes an XML document entry to the archive.
     *
     * @param ZipArchive $archive The archive.
     * @param string $path The entry path.
     * @param XDocument $document The document.
     * @return void
     */
    public static function WriteXmlEntry(ZipArchive $archive, $path, XDocument $document)
    {
        $archive->addEntry($path, $document->save());
    }

    /**
     * Builds the [Content_Types].xml document.
     *
     * @param WorkbookModel $model The workbook model.
     * @param bool $hasSharedStrings Whether shared strings exist.
     * @param bool $hasDateStyles Whether styles exist.
     * @param bool $hasCoreProperties Whether core properties exist.
     * @param bool $hasExtendedProperties Whether extended properties exist.
     * @param int $totalTableCount The total table count.
     * @param int $totalDrawingCount The total drawing count.
     * @param array $imageExtensions The image extensions.
     * @param array $chartPartNames The chart part names.
     * @param array $chartContentTypes The chart content types.
     * @param array $chartCompanionPartNames The chart companion part names.
     * @param array $chartCompanionContentTypes The chart companion content types.
     * @param bool $hasTheme Whether a theme part exists.
     * @param int $totalCommentCount The total comment count.
     * @return XDocument The document.
     */
    public static function BuildContentTypes(WorkbookModel $model, $hasSharedStrings, $hasDateStyles, $hasCoreProperties, $hasExtendedProperties, $totalTableCount, $totalDrawingCount, array $imageExtensions, array $chartPartNames, array $chartContentTypes, array $chartCompanionPartNames, array $chartCompanionContentTypes, $hasTheme = false, $totalCommentCount = 0)
    {
        $root = new XElement(
            self::ContentTypeNs . 'Types',
            new XElement(
                self::ContentTypeNs . 'Default',
                new XAttribute('Extension', 'rels'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-package.relationships+xml')
            ),
            new XElement(
                self::ContentTypeNs . 'Default',
                new XAttribute('Extension', 'xml'),
                new XAttribute('ContentType', 'application/xml')
            ),
            new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/workbook.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
            )
        );

        if ($hasCoreProperties) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/docProps/core.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml')
            ));
        }

        if ($hasExtendedProperties) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/docProps/app.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')
            ));
        }

        for ($index = 0; $index < count($model->Worksheets); $index++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/worksheets/sheet' . ($index + 1) . '.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
            ));
        }

        for ($tableNumber = 1; $tableNumber <= $totalTableCount; $tableNumber++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/tables/table' . $tableNumber . '.xml'),
                new XAttribute('ContentType', XlsxWorkbookTables::TableContentType)
            ));
        }

        for ($drawingNumber = 1; $drawingNumber <= $totalDrawingCount; $drawingNumber++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/drawings/drawing' . $drawingNumber . '.xml'),
                new XAttribute('ContentType', XlsxWorkbookPictures::DrawingContentType)
            ));
        }

        for ($c = 0; $c < count($chartPartNames); $c++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', $chartPartNames[$c]),
                new XAttribute('ContentType', $chartContentTypes[$c])
            ));
        }

        for ($c = 0; $c < count($chartCompanionPartNames); $c++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', $chartCompanionPartNames[$c]),
                new XAttribute('ContentType', $chartCompanionContentTypes[$c])
            ));
        }

        for ($extIndex = 0; $extIndex < count($imageExtensions); $extIndex++) {
            $ext = $imageExtensions[$extIndex];
            $root->Add(new XElement(
                self::ContentTypeNs . 'Default',
                new XAttribute('Extension', $ext),
                new XAttribute('ContentType', Picture::ContentTypeFromExtension($ext))
            ));
        }

        if ($hasSharedStrings) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/sharedStrings.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
            ));
        }

        if ($hasDateStyles) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/styles.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
            ));
        }

        if ($hasTheme) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/theme/theme1.xml'),
                new XAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml')
            ));
        }

        for ($n = 1; $n <= $totalCommentCount; $n++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/comments' . $n . '.xml'),
                new XAttribute('ContentType', XlsxWorkbookComments::CommentsContentType)
            ));
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/drawings/vmlDrawing' . $n . '.vml'),
                new XAttribute('ContentType', XlsxWorkbookComments::VmlDrawingContentType)
            ));
        }

        for ($i = 0; $i < count($model->ExternalLinks); $i++) {
            $root->Add(new XElement(
                self::ContentTypeNs . 'Override',
                new XAttribute('PartName', '/xl/externalLinks/externalLink' . ($i + 1) . '.xml'),
                new XAttribute('ContentType', self::ExternalLinkContentType)
            ));
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $root);
    }

    /**
     * Builds the package root relationships document.
     *
     * @param bool $hasCoreProperties Whether core properties exist.
     * @param bool $hasExtendedProperties Whether extended properties exist.
     * @return XDocument The document.
     */
    public static function BuildRootRelationships($hasCoreProperties, $hasExtendedProperties)
    {
        $relationships = new XElement(
            self::PackageRelationshipNs . 'Relationships',
            new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId1'),
                new XAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'),
                new XAttribute('Target', 'xl/workbook.xml')
            )
        );

        $relationshipId = 2;
        if ($hasCoreProperties) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', self::CorePropertiesRelationshipType),
                new XAttribute('Target', 'docProps/core.xml')
            ));
            $relationshipId++;
        }

        if ($hasExtendedProperties) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', self::ExtendedPropertiesRelationshipType),
                new XAttribute('Target', 'docProps/app.xml')
            ));
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $relationships);
    }

    /**
     * Builds the workbook.xml document.
     *
     * @param WorkbookModel $model The workbook model.
     * @param int $externalLinkBaseRId The external link base relationship id.
     * @return XDocument The document.
     */
    public static function BuildWorkbook(WorkbookModel $model, $externalLinkBaseRId)
    {
        $workbook = new XElement(
            self::MainNs . 'workbook',
            new XAttribute(XNamespace::Xmlns . 'r', self::RelationshipNs)
        );

        $workbookProperties = XlsxWorkbookProperties::BuildWorkbookPropertiesElement($model);
        if ($workbookProperties !== null) {
            $workbook->Add($workbookProperties);
        }

        $workbookProtection = XlsxWorkbookProperties::BuildWorkbookProtectionElement($model);
        if ($workbookProtection !== null) {
            $workbook->Add($workbookProtection);
        }

        $bookViews = XlsxWorkbookProperties::BuildBookViewsElement($model);
        if ($bookViews !== null) {
            $workbook->Add($bookViews);
        }

        $sheets = new XElement(self::MainNs . 'sheets');
        for ($index = 0; $index < count($model->Worksheets); $index++) {
            $worksheet = $model->Worksheets[$index];
            $sheet = new XElement(
                self::MainNs . 'sheet',
                new XAttribute('name', $worksheet->Name),
                new XAttribute('sheetId', $index + 1),
                new XAttribute(self::RelationshipNs . 'id', 'rId' . ($index + 1))
            );

            if ($worksheet->Visibility === SheetVisibility::Hidden) {
                $sheet->SetAttributeValue('state', 'hidden');
            } elseif ($worksheet->Visibility === SheetVisibility::VeryHidden) {
                $sheet->SetAttributeValue('state', 'veryHidden');
            }

            $sheets->Add($sheet);
        }

        $workbook->Add($sheets);

        if (count($model->ExternalLinks) > 0) {
            $externalReferences = new XElement(self::MainNs . 'externalReferences');
            for ($i = 0; $i < count($model->ExternalLinks); $i++) {
                $externalReferences->Add(new XElement(
                    self::MainNs . 'externalReference',
                    new XAttribute(self::RelationshipNs . 'id', 'rId' . ($externalLinkBaseRId + $i))
                ));
            }

            $workbook->Add($externalReferences);
        }

        $definedNames = XlsxWorkbookDefinedNames::BuildDefinedNames($model);
        if ($definedNames !== null) {
            $workbook->Add($definedNames);
        }

        $calculationProperties = XlsxWorkbookProperties::BuildCalculationPropertiesElement($model);
        if ($calculationProperties !== null) {
            $workbook->Add($calculationProperties);
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $workbook);
    }

    /**
     * Builds the workbook relationships document.
     *
     * @param WorkbookModel $model The workbook model.
     * @param bool $hasSharedStrings Whether shared strings exist.
     * @param bool $hasDateStyles Whether styles exist.
     * @return XDocument The document.
     */
    public static function BuildWorkbookRelationships(WorkbookModel $model, $hasSharedStrings, $hasDateStyles)
    {
        $relationships = new XElement(self::PackageRelationshipNs . 'Relationships');
        $relationshipId = 1;

        for ($index = 0; $index < count($model->Worksheets); $index++) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', self::WorksheetRelationshipType),
                new XAttribute('Target', 'worksheets/sheet' . ($index + 1) . '.xml')
            ));
            $relationshipId++;
        }

        if ($hasSharedStrings) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', self::SharedStringsRelationshipType),
                new XAttribute('Target', 'sharedStrings.xml')
            ));
            $relationshipId++;
        }

        if ($hasDateStyles) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', self::StylesRelationshipType),
                new XAttribute('Target', 'styles.xml')
            ));
            $relationshipId++;
        }

        if (!($model->RawThemeXml === null || $model->RawThemeXml === '')) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'),
                new XAttribute('Target', 'theme/theme1.xml')
            ));
            $relationshipId++;
        }

        for ($i = 0; $i < count($model->ExternalLinks); $i++) {
            $relationships->Add(new XElement(
                self::PackageRelationshipNs . 'Relationship',
                new XAttribute('Id', 'rId' . $relationshipId),
                new XAttribute('Type', self::ExternalLinkRelationshipType),
                new XAttribute('Target', 'externalLinks/externalLink' . ($i + 1) . '.xml')
            ));
            $relationshipId++;
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $relationships);
    }

    /**
     * Converts a merge region to an A1 range reference.
     *
     * @param MergeRegion $region The region.
     * @return string The range reference.
     */
    public static function ToRangeReference(MergeRegion $region)
    {
        $first = (new CellAddress($region->FirstRow, $region->FirstColumn))->ToString();
        $last = (new CellAddress($region->FirstRow + $region->TotalRows - 1, $region->FirstColumn + $region->TotalColumns - 1))->ToString();
        return $first === $last ? $first : $first . ':' . $last;
    }

    /**
     * Attempts to parse a merge reference.
     *
     * @param string $mergeReference The merge reference.
     * @param MergeRegion|null $region The parsed region.
     * @return bool True when parsing succeeds.
     */
    public static function TryParseMergeReference($mergeReference, &$region)
    {
        $region = null;
        if ($mergeReference === null || trim($mergeReference) === '') {
            return false;
        }

        $parts = explode(':', $mergeReference);
        if (count($parts) === 1) {
            $single = null;
            if (!self::TryParseCellReference($parts[0], $single)) {
                return false;
            }

            $region = new MergeRegion($single->RowIndex, $single->ColumnIndex, 1, 1);
            return true;
        }

        $first = null;
        $last = null;
        if (count($parts) !== 2
            || !self::TryParseCellReference($parts[0], $first)
            || !self::TryParseCellReference($parts[1], $last)
            || $last->RowIndex < $first->RowIndex
            || $last->ColumnIndex < $first->ColumnIndex) {
            return false;
        }

        $region = new MergeRegion($first->RowIndex, $first->ColumnIndex, $last->RowIndex - $first->RowIndex + 1, $last->ColumnIndex - $first->ColumnIndex + 1);
        return true;
    }

    /**
     * Attempts to parse a cell reference.
     *
     * @param string $cellReference The cell reference.
     * @param CellAddress|null $address The parsed address.
     * @return bool True when parsing succeeds.
     */
    public static function TryParseCellReference($cellReference, &$address)
    {
        try {
            $address = CellAddress::Parse($cellReference);
            return true;
        } catch (\InvalidArgumentException $exception) {
            $address = null;
            return false;
        }
    }

    /**
     * Builds the shared strings document.
     *
     * @param SharedStringRepository $sharedStrings The shared strings.
     * @return XDocument The document.
     */
    public static function BuildSharedStrings(SharedStringRepository $sharedStrings)
    {
        $root = new XElement(
            self::MainNs . 'sst',
            new XAttribute('count', count($sharedStrings->Values)),
            new XAttribute('uniqueCount', count($sharedStrings->Values))
        );

        foreach ($sharedStrings->Values as $value) {
            $root->Add(new XElement(self::MainNs . 'si', self::CreateTextElement($value)));
        }

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $root);
    }

    /**
     * Creates an inline text element.
     *
     * @param string $value The text.
     * @return XElement The element.
     */
    public static function CreateTextElement($value)
    {
        $text = new XElement(self::MainNs . 't', $value);
        if (self::NeedsPreserveWhitespace($value)) {
            $text->SetAttributeValue(self::XmlNs . 'space', 'preserve');
        }

        return $text;
    }

    /**
     * Determines whether text requires xml:space="preserve".
     *
     * @param string $value The text.
     * @return bool True when whitespace must be preserved.
     */
    public static function NeedsPreserveWhitespace($value)
    {
        if ($value === null || $value === '') {
            return false;
        }

        $length = strlen($value);
        return self::IsWhiteSpace($value[0])
            || self::IsWhiteSpace($value[$length - 1])
            || strpos($value, "\n") !== false
            || strpos($value, "\r") !== false
            || strpos($value, "\t") !== false;
    }

    /**
     * Determines whether a character is whitespace.
     *
     * @param string $character The character.
     * @return bool True when the character is whitespace.
     */
    private static function IsWhiteSpace($character)
    {
        return $character === ' ' || $character === "\t" || $character === "\n" || $character === "\r";
    }

    /**
     * Builds a minimal stylesheet document.
     *
     * @return XDocument The document.
     */
    public static function BuildMinimalStylesheet()
    {
        $stylesheet = new XElement(
            self::MainNs . 'styleSheet',
            new XElement(
                self::MainNs . 'fonts',
                new XAttribute('count', 1),
                new XElement(
                    self::MainNs . 'font',
                    new XElement(self::MainNs . 'sz', new XAttribute('val', 11)),
                    new XElement(self::MainNs . 'name', new XAttribute('val', 'Calibri'))
                )
            ),
            new XElement(
                self::MainNs . 'fills',
                new XAttribute('count', 2),
                new XElement(self::MainNs . 'fill', new XElement(self::MainNs . 'patternFill', new XAttribute('patternType', 'none'))),
                new XElement(self::MainNs . 'fill', new XElement(self::MainNs . 'patternFill', new XAttribute('patternType', 'gray125')))
            ),
            new XElement(
                self::MainNs . 'borders',
                new XAttribute('count', 1),
                new XElement(
                    self::MainNs . 'border',
                    new XElement(self::MainNs . 'left'),
                    new XElement(self::MainNs . 'right'),
                    new XElement(self::MainNs . 'top'),
                    new XElement(self::MainNs . 'bottom'),
                    new XElement(self::MainNs . 'diagonal')
                )
            ),
            new XElement(
                self::MainNs . 'cellStyleXfs',
                new XAttribute('count', 1),
                new XElement(
                    self::MainNs . 'xf',
                    new XAttribute('numFmtId', 0),
                    new XAttribute('fontId', 0),
                    new XAttribute('fillId', 0),
                    new XAttribute('borderId', 0)
                )
            ),
            new XElement(
                self::MainNs . 'cellXfs',
                new XAttribute('count', 2),
                new XElement(
                    self::MainNs . 'xf',
                    new XAttribute('numFmtId', 0),
                    new XAttribute('fontId', 0),
                    new XAttribute('fillId', 0),
                    new XAttribute('borderId', 0),
                    new XAttribute('xfId', 0)
                ),
                new XElement(
                    self::MainNs . 'xf',
                    new XAttribute('numFmtId', 14),
                    new XAttribute('fontId', 0),
                    new XAttribute('fillId', 0),
                    new XAttribute('borderId', 0),
                    new XAttribute('xfId', 0),
                    new XAttribute('applyNumberFormat', 1)
                )
            ),
            new XElement(
                self::MainNs . 'cellStyles',
                new XAttribute('count', 1),
                new XElement(
                    self::MainNs . 'cellStyle',
                    new XAttribute('name', 'Normal'),
                    new XAttribute('xfId', 0),
                    new XAttribute('builtinId', 0)
                )
            )
        );

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $stylesheet);
    }
}
