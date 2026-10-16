<?php

namespace Aspose\Cells_FOSS\Packaging;

final class PackagingConventions
{
    const ContentTypesPart = '/[Content_Types].xml';
    const RootRelationshipsPart = '/_rels/.rels';
    const WorkbookPart = '/xl/workbook.xml';
    const WorkbookRelationshipsPart = '/xl/_rels/workbook.xml.rels';
    const OfficeDocumentRelationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    const WorksheetRelationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    const StylesRelationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
    const SharedStringsRelationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';

    private function __construct()
    {
    }
}
