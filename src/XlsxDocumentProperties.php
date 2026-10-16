<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDeclaration;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XmlName;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

final class XlsxDocumentProperties
{
    const CorePropertiesNs = '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}';
    const DublinCoreNs = '{http://purl.org/dc/elements/1.1/}';
    const DublinCoreTermsNs = '{http://purl.org/dc/terms/}';
    const XsiNs = '{http://www.w3.org/2001/XMLSchema-instance}';
    const ExtendedPropertiesNs = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}';

    public static function BuildCorePropertiesDocument(WorkbookModel $model)
    {
        if (!$model->DocumentProperties->Core->HasStoredState()) {
            return null;
        }

        $core = $model->DocumentProperties->Core;
        $root = new XElement(self::CorePropertiesNs . 'coreProperties',
            new XAttribute(XNamespace::Xmlns . 'cp', self::CorePropertiesNs),
            new XAttribute(XNamespace::Xmlns . 'dc', self::DublinCoreNs),
            new XAttribute(XNamespace::Xmlns . 'dcterms', self::DublinCoreTermsNs),
            new XAttribute(XNamespace::Xmlns . 'xsi', self::XsiNs));

        self::AddStringElement($root, self::DublinCoreNs . 'title', $core->Title);
        self::AddStringElement($root, self::DublinCoreNs . 'subject', $core->Subject);
        self::AddStringElement($root, self::DublinCoreNs . 'creator', $core->Creator);
        self::AddStringElement($root, self::CorePropertiesNs . 'keywords', $core->Keywords);
        self::AddStringElement($root, self::DublinCoreNs . 'description', $core->Description);
        self::AddStringElement($root, self::CorePropertiesNs . 'lastModifiedBy', $core->LastModifiedBy);
        self::AddStringElement($root, self::CorePropertiesNs . 'revision', $core->Revision);
        self::AddStringElement($root, self::CorePropertiesNs . 'category', $core->Category);
        self::AddStringElement($root, self::CorePropertiesNs . 'contentStatus', $core->ContentStatus);
        self::AddDateElement($root, self::DublinCoreTermsNs . 'created', $core->Created);
        self::AddDateElement($root, self::DublinCoreTermsNs . 'modified', $core->Modified);

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $root);
    }

    public static function BuildExtendedPropertiesDocument(WorkbookModel $model)
    {
        if (!$model->DocumentProperties->Extended->HasStoredState()) {
            return null;
        }

        $extended = $model->DocumentProperties->Extended;
        $root = new XElement(self::ExtendedPropertiesNs . 'Properties',
            new XAttribute(XNamespace::Xmlns . 'vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'));

        self::AddStringElement($root, self::ExtendedPropertiesNs . 'Application', $extended->Application);
        self::AddStringElement($root, self::ExtendedPropertiesNs . 'AppVersion', $extended->AppVersion);
        self::AddStringElement($root, self::ExtendedPropertiesNs . 'Company', $extended->Company);
        self::AddStringElement($root, self::ExtendedPropertiesNs . 'Manager', $extended->Manager);
        self::AddIntElement($root, self::ExtendedPropertiesNs . 'DocSecurity', $extended->DocSecurity);
        self::AddStringElement($root, self::ExtendedPropertiesNs . 'HyperlinkBase', $extended->HyperlinkBase);
        self::AddBoolElement($root, self::ExtendedPropertiesNs . 'ScaleCrop', $extended->ScaleCrop);
        self::AddBoolElement($root, self::ExtendedPropertiesNs . 'LinksUpToDate', $extended->LinksUpToDate);
        self::AddBoolElement($root, self::ExtendedPropertiesNs . 'SharedDoc', $extended->SharedDoc);

        return new XDocument(new XDeclaration('1.0', 'utf-8', 'yes'), $root);
    }

    public static function LoadDocumentProperties(ZipArchive $archive, WorkbookModel $model, LoadDiagnostics $diagnostics, LoadOptions $options)
    {
        $relationshipTargets = self::LoadRootDocumentPropertiesRelationshipTargets($archive, $diagnostics, $options);
        $corePropertiesPartUri = '';
        if (array_key_exists(XlsxWorkbookSerializerCommon::CorePropertiesRelationshipType, $relationshipTargets)) {
            $resolvedCorePropertiesPartUri = $relationshipTargets[XlsxWorkbookSerializerCommon::CorePropertiesRelationshipType];
            if (!($resolvedCorePropertiesPartUri === null || $resolvedCorePropertiesPartUri === '')) {
                $corePropertiesPartUri = $resolvedCorePropertiesPartUri;
            }
        }

        $extendedPropertiesPartUri = '';
        if (array_key_exists(XlsxWorkbookSerializerCommon::ExtendedPropertiesRelationshipType, $relationshipTargets)) {
            $resolvedExtendedPropertiesPartUri = $relationshipTargets[XlsxWorkbookSerializerCommon::ExtendedPropertiesRelationshipType];
            if (!($resolvedExtendedPropertiesPartUri === null || $resolvedExtendedPropertiesPartUri === '')) {
                $extendedPropertiesPartUri = $resolvedExtendedPropertiesPartUri;
            }
        }

        self::LoadCoreProperties($archive, $model, $diagnostics, $options, $corePropertiesPartUri);
        self::LoadExtendedProperties($archive, $model, $diagnostics, $options, $extendedPropertiesPartUri);
    }

    private static function LoadRootDocumentPropertiesRelationshipTargets(ZipArchive $archive, LoadDiagnostics $diagnostics, LoadOptions $options)
    {
        $relationshipTargets = array();
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, '/_rels/.rels');
        if ($entry === null) {
            return $relationshipTargets;
        }

        try {
            $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        } catch (\Throwable $exception) {
            self::AddDocumentPropertiesIssue($diagnostics, $options, '/_rels/.rels', 'Package root relationships were malformed and document properties were ignored.');
            return $relationshipTargets;
        }

        $root = $document->Root;
        $elements = $root === null ? array() : $root->Elements(XlsxWorkbookSerializerCommon::PackageRelationshipNs . 'Relationship');
        foreach ($elements as $relationship) {
            $type = (string)$relationship->Attribute('Type');
            $target = (string)$relationship->Attribute('Target');
            $targetMode = (string)$relationship->Attribute('TargetMode');
            if ($type === '' || $target === '' || strcasecmp($targetMode, 'External') === 0) {
                continue;
            }

            $relationshipType = $type;
            if (strcasecmp($relationshipType, XlsxWorkbookSerializerCommon::CorePropertiesRelationshipType) !== 0
                && strcasecmp($relationshipType, XlsxWorkbookSerializerCommon::ExtendedPropertiesRelationshipType) !== 0) {
                continue;
            }

            if (!array_key_exists($relationshipType, $relationshipTargets)) {
                $relationshipTargets[$relationshipType] = XlsxWorkbookArchiveHelpers::ResolvePartUri('/', $target);
            }
        }

        return $relationshipTargets;
    }

    private static function LoadCoreProperties(ZipArchive $archive, WorkbookModel $model, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        if ($partUri === null || $partUri === '') {
            return;
        }

        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $partUri);
        if ($entry === null) {
            return;
        }

        try {
            $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        } catch (\Throwable $exception) {
            self::AddDocumentPropertiesIssue($diagnostics, $options, $partUri, 'Core document properties part was malformed and was ignored.');
            return;
        }

        $root = $document->Root;
        if ($root === null) {
            return;
        }

        $core = $model->DocumentProperties->Core;
        $core->Title = self::ReadElementValue($root, self::DublinCoreNs . 'title');
        $core->Subject = self::ReadElementValue($root, self::DublinCoreNs . 'subject');
        $core->Creator = self::ReadElementValue($root, self::DublinCoreNs . 'creator');
        $core->Keywords = self::ReadElementValue($root, self::CorePropertiesNs . 'keywords');
        $core->Description = self::ReadElementValue($root, self::DublinCoreNs . 'description');
        $core->LastModifiedBy = self::ReadElementValue($root, self::CorePropertiesNs . 'lastModifiedBy');
        $core->Revision = self::ReadElementValue($root, self::CorePropertiesNs . 'revision');
        $core->Category = self::ReadElementValue($root, self::CorePropertiesNs . 'category');
        $core->ContentStatus = self::ReadElementValue($root, self::CorePropertiesNs . 'contentStatus');
        $core->Created = self::ReadDateElement($root, self::DublinCoreTermsNs . 'created', $diagnostics, $options, $partUri);
        $core->Modified = self::ReadDateElement($root, self::DublinCoreTermsNs . 'modified', $diagnostics, $options, $partUri);
    }

    private static function LoadExtendedProperties(ZipArchive $archive, WorkbookModel $model, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        if ($partUri === null || $partUri === '') {
            return;
        }

        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $partUri);
        if ($entry === null) {
            return;
        }

        try {
            $document = XlsxWorkbookArchiveHelpers::LoadDocument($entry);
        } catch (\Throwable $exception) {
            self::AddDocumentPropertiesIssue($diagnostics, $options, $partUri, 'Extended document properties part was malformed and was ignored.');
            return;
        }

        $root = $document->Root;
        if ($root === null) {
            return;
        }

        $extended = $model->DocumentProperties->Extended;
        $extended->Application = self::ReadElementValue($root, self::ExtendedPropertiesNs . 'Application');
        $extended->AppVersion = self::ReadElementValue($root, self::ExtendedPropertiesNs . 'AppVersion');
        $extended->Company = self::ReadElementValue($root, self::ExtendedPropertiesNs . 'Company');
        $extended->Manager = self::ReadElementValue($root, self::ExtendedPropertiesNs . 'Manager');
        $extended->DocSecurity = self::ReadIntElement($root, self::ExtendedPropertiesNs . 'DocSecurity', $diagnostics, $options, $partUri);
        $extended->HyperlinkBase = self::ReadElementValue($root, self::ExtendedPropertiesNs . 'HyperlinkBase');
        $extended->ScaleCrop = self::ReadBoolElement($root, self::ExtendedPropertiesNs . 'ScaleCrop', $diagnostics, $options, $partUri);
        $extended->LinksUpToDate = self::ReadBoolElement($root, self::ExtendedPropertiesNs . 'LinksUpToDate', $diagnostics, $options, $partUri);
        $extended->SharedDoc = self::ReadBoolElement($root, self::ExtendedPropertiesNs . 'SharedDoc', $diagnostics, $options, $partUri);
    }

    private static function AddStringElement($parent, $name, $value)
    {
        if (!($value === null || $value === '')) {
            $parent->Add(new XElement($name, $value));
        }
    }

    private static function AddIntElement($parent, $name, $value)
    {
        if ($value !== null) {
            $parent->Add(new XElement($name, (string)$value));
        }
    }

    private static function AddBoolElement($parent, $name, $value)
    {
        if ($value !== null) {
            $parent->Add(new XElement($name, $value ? 'true' : 'false'));
        }
    }

    private static function AddDateElement($parent, $name, $value)
    {
        if ($value === null) {
            return;
        }

        $element = new XElement($name, $value->format('c'));
        $element->SetAttributeValue(self::XsiNs . 'type', 'dcterms:W3CDTF');
        $parent->Add($element);
    }

    private static function ReadElementValue($parent, $name)
    {
        $element = $parent->Element($name);
        $value = $element === null ? '' : (string)$element->Value;
        return trim($value);
    }

    private static function ReadDateElement($parent, $name, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        $element = $parent->Element($name);
        if ($element === null) {
            return null;
        }

        $localName = XmlName::split($name)[1];
        try {
            return new \DateTime((string)$element->Value);
        } catch (\Throwable $exception) {
            self::AddDocumentPropertiesIssue($diagnostics, $options, $partUri, "Document property '" . $localName . "' had an invalid timestamp and was ignored.");
            return null;
        }
    }

    private static function ReadIntElement($parent, $name, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        $element = $parent->Element($name);
        if ($element === null) {
            return null;
        }

        $localName = XmlName::split($name)[1];
        $rawValue = trim((string)$element->Value);
        if (preg_match('/^[+-]?[0-9]+$/', $rawValue) === 1) {
            $value = intval($rawValue);
            if ($value >= 0) {
                return $value;
            }
        }

        self::AddDocumentPropertiesIssue($diagnostics, $options, $partUri, "Document property '" . $localName . "' had an invalid integer value and was ignored.");
        return null;
    }

    private static function ReadBoolElement($parent, $name, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        $element = $parent->Element($name);
        if ($element === null) {
            return null;
        }

        $localName = XmlName::split($name)[1];
        $rawValue = (string)$element->Value;
        if ($rawValue === '1' || strcasecmp($rawValue, 'true') === 0) {
            return true;
        }

        if ($rawValue === '0' || strcasecmp($rawValue, 'false') === 0) {
            return false;
        }

        self::AddDocumentPropertiesIssue($diagnostics, $options, $partUri, "Document property '" . $localName . "' had an invalid Boolean value and was ignored.");
        return null;
    }

    private static function AddDocumentPropertiesIssue(LoadDiagnostics $diagnostics, LoadOptions $options, $partUri, $message)
    {
        $issue = new LoadIssue('WB-L004', DiagnosticSeverity::Warning, $message);
        $issue->PartUri = $partUri;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
    }
}
