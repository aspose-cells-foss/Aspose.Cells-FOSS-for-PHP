<?php

namespace Aspose\Cells_FOSS\PageSetupOpenXmlTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';

use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\PageSetupScenarioFactory;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\ZipPackageHelper;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    public static function Main()
    {
        return TestRunner::Run(
            'PageSetup.OpenXmlTests',
            array(
                new TestCase('page_setup_matches_openxml_sdk', array(Program::class, 'PageSetupMatchesOpenXmlSdk')),
                new TestCase('openxml_created_page_setup_loads_in_library', array(Program::class, 'OpenXmlCreatedPageSetupLoadsInLibrary')),
            )
        );
    }

    public static function PageSetupMatchesOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('page-setup-openxml-write');
        $path = $temp->GetPath('page-setup.xlsx');

        $workbook = PageSetupScenarioFactory::CreatePageSetupWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        PageSetupScenarioFactory::AssertPageSetup($loaded);

        $snapshot = self::ReadPageSetupSnapshot($path, 'Print Sheet');
        AssertEx::Equal('$A$1:$C$10', $snapshot['PrintArea']);
        AssertEx::Equal('$1:$2', $snapshot['PrintTitleRows']);
        AssertEx::Equal('$A:$B', $snapshot['PrintTitleColumns']);
        AssertEx::Equal(0.25, $snapshot['LeftMargin'] ?? 0.0);
        AssertEx::Equal(0.4, $snapshot['RightMargin'] ?? 0.0);
        AssertEx::Equal(0.5, $snapshot['TopMargin'] ?? 0.0);
        AssertEx::Equal(0.6, $snapshot['BottomMargin'] ?? 0.0);
        AssertEx::Equal(0.2, $snapshot['HeaderMargin'] ?? 0.0);
        AssertEx::Equal(0.22, $snapshot['FooterMargin'] ?? 0.0);
        AssertEx::Equal(9, $snapshot['PaperSize'] ?? 0);
        AssertEx::Equal('landscape', $snapshot['Orientation']);
        AssertEx::Equal(3, $snapshot['FirstPageNumber'] ?? 0);
        AssertEx::Equal(95, $snapshot['Scale'] ?? 0);
        AssertEx::Equal(1, $snapshot['FitToWidth'] ?? 0);
        AssertEx::Equal(2, $snapshot['FitToHeight'] ?? 0);
        AssertEx::True($snapshot['PrintGridLines']);
        AssertEx::True($snapshot['PrintHeadings']);
        AssertEx::True($snapshot['HorizontalCentered']);
        AssertEx::True($snapshot['VerticalCentered']);
        AssertEx::Equal('&LLeft Header&CCenter Header&RRight Header', $snapshot['OddHeader']);
        AssertEx::Equal('&LLeft Footer&CCenter Footer&RRight Footer', $snapshot['OddFooter']);
        AssertEx::Equal(2, count($snapshot['RowBreaks']));
        AssertEx::Equal(4, $snapshot['RowBreaks'][0]);
        AssertEx::Equal(7, $snapshot['RowBreaks'][1]);
        AssertEx::Equal(1, count($snapshot['ColumnBreaks']));
        AssertEx::Equal(2, $snapshot['ColumnBreaks'][0]);
    }

    public static function OpenXmlCreatedPageSetupLoadsInLibrary()
    {
        $temp = new TemporaryDirectory('page-setup-openxml-read');
        $path = $temp->GetPath('openxml-page-setup.xlsx');

        self::CreateOpenXmlPageSetupWorkbook($path);

        $loaded = new Workbook($path);
        PageSetupScenarioFactory::AssertPageSetup($loaded);

        $snapshot = self::ReadPageSetupSnapshot($path, 'Print Sheet');
        AssertEx::Equal('$A$1:$C$10', $snapshot['PrintArea']);
        AssertEx::Equal('$1:$2', $snapshot['PrintTitleRows']);
        AssertEx::Equal('$A:$B', $snapshot['PrintTitleColumns']);
    }

    private static function ReadPageSetupSnapshot($workbookPath, $sheetName)
    {
        $archive = new \ZipArchive();
        if ($archive->open($workbookPath) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $workbookXml = $archive->getFromName('xl/workbook.xml');
        $relationshipsXml = $archive->getFromName('xl/_rels/workbook.xml.rels');
        $archive->close();

        if ($workbookXml === false) {
            throw new \RuntimeException('WorkbookPart is missing.');
        }

        $workbookDocument = XDocument::Load($workbookXml);
        $relationships = array();
        if ($relationshipsXml !== false) {
            $relationshipsDocument = XDocument::Load($relationshipsXml);
            if ($relationshipsDocument->Root !== null) {
                foreach ($relationshipsDocument->Root->Elements('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship') as $relationship) {
                    $id = (string)$relationship->Attribute('Id');
                    $target = (string)$relationship->Attribute('Target');
                    if ($id !== '' && $target !== '') {
                        $relationships[$id] = self::ResolvePartUri('/xl/workbook.xml', $target);
                    }
                }
            }
        }

        $sheetElement = null;
        if ($workbookDocument->Root !== null) {
            foreach ($workbookDocument->Root->Elements(self::MainNs . 'sheets') as $sheetsElement) {
                foreach ($sheetsElement->Elements(self::MainNs . 'sheet') as $candidate) {
                    if ((string)$candidate->Attribute('name') === $sheetName) {
                        $sheetElement = $candidate;
                        break 2;
                    }
                }
            }
        }

        if ($sheetElement === null) {
            throw new \RuntimeException('Worksheet was not found.');
        }

        $sheetIdAttribute = $sheetElement->Attribute('sheetId');
        $sheetId = intval($sheetIdAttribute === null ? '1' : (string)$sheetIdAttribute->Value) - 1;
        $relationshipId = (string)$sheetElement->Attribute(self::RelationshipNs . 'id');
        if ($relationshipId === '') {
            throw new \RuntimeException('Worksheet relationship id is missing.');
        }

        if (!array_key_exists($relationshipId, $relationships)) {
            throw new \RuntimeException('Worksheet relationship was not found.');
        }

        $worksheetUri = ltrim($relationships[$relationshipId], '/');

        $archive = new \ZipArchive();
        if ($archive->open($workbookPath) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $worksheetXml = $archive->getFromName($worksheetUri);
        $archive->close();
        if ($worksheetXml === false) {
            throw new \RuntimeException('Worksheet part was not found.');
        }

        $worksheetDocument = XDocument::Load($worksheetXml);
        $root = $worksheetDocument->Root;
        if ($root === null) {
            throw new \RuntimeException('Worksheet XML is empty.');
        }

        $margins = $root->Element(self::MainNs . 'pageMargins');
        $pageSetup = $root->Element(self::MainNs . 'pageSetup');
        $printOptions = $root->Element(self::MainNs . 'printOptions');
        $headerFooter = $root->Element(self::MainNs . 'headerFooter');

        $rowBreaks = array();
        $rowBreaksElement = $root->Element(self::MainNs . 'rowBreaks');
        if ($rowBreaksElement !== null) {
            foreach ($rowBreaksElement->Elements(self::MainNs . 'brk') as $element) {
                $rowBreaks[] = self::AttributeIntOrZero($element, 'id');
            }
        }

        $columnBreaks = array();
        $colBreaksElement = $root->Element(self::MainNs . 'colBreaks');
        if ($colBreaksElement !== null) {
            foreach ($colBreaksElement->Elements(self::MainNs . 'brk') as $element) {
                $columnBreaks[] = self::AttributeIntOrZero($element, 'id');
            }
        }

        $printArea = null;
        $printTitleRows = null;
        $printTitleColumns = null;
        if ($workbookDocument->Root !== null) {
            foreach ($workbookDocument->Root->Elements(self::MainNs . 'definedNames') as $definedNamesElement) {
                foreach ($definedNamesElement->Elements(self::MainNs . 'definedName') as $definedName) {
                    $localSheetIdAttribute = $definedName->Attribute('localSheetId');
                    $localSheetId = $localSheetIdAttribute === null ? '' : (string)$localSheetIdAttribute->Value;
                    if ($localSheetId !== (string)$sheetId) {
                        continue;
                    }

                    $nameAttribute = $definedName->Attribute('name');
                    $name = $nameAttribute === null ? '' : (string)$nameAttribute->Value;
                    if ($name === '_xlnm.Print_Area') {
                        $printArea = self::UnqualifyDefinedNameValue($definedName->Value);
                    } elseif ($name === '_xlnm.Print_Titles') {
                        foreach (explode(',', $definedName->Value) as $segment) {
                            $unqualified = self::UnqualifyDefinedNameValue($segment);
                            $stripped = str_replace('$', '', $unqualified);
                            if (strpos($unqualified, '$') !== false
                                && preg_match('/[0-9]/', $unqualified) === 1
                                && strpos($unqualified, ':') !== false
                                && $stripped !== ''
                                && ctype_digit($stripped[0])) {
                                $printTitleRows = $unqualified;
                            } else {
                                $printTitleColumns = $unqualified;
                            }
                        }
                    }
                }
            }
        }

        return array(
            'PrintArea' => $printArea ?? '',
            'PrintTitleRows' => $printTitleRows ?? '',
            'PrintTitleColumns' => $printTitleColumns ?? '',
            'LeftMargin' => self::TryParseDouble($margins === null ? null : self::AttributeValue($margins, 'left')),
            'RightMargin' => self::TryParseDouble($margins === null ? null : self::AttributeValue($margins, 'right')),
            'TopMargin' => self::TryParseDouble($margins === null ? null : self::AttributeValue($margins, 'top')),
            'BottomMargin' => self::TryParseDouble($margins === null ? null : self::AttributeValue($margins, 'bottom')),
            'HeaderMargin' => self::TryParseDouble($margins === null ? null : self::AttributeValue($margins, 'header')),
            'FooterMargin' => self::TryParseDouble($margins === null ? null : self::AttributeValue($margins, 'footer')),
            'PaperSize' => self::TryParseInt($pageSetup === null ? null : self::AttributeValue($pageSetup, 'paperSize')),
            'Orientation' => $pageSetup === null ? '' : (self::AttributeValue($pageSetup, 'orientation') ?? ''),
            'FirstPageNumber' => self::TryParseInt($pageSetup === null ? null : self::AttributeValue($pageSetup, 'firstPageNumber')),
            'Scale' => self::TryParseInt($pageSetup === null ? null : self::AttributeValue($pageSetup, 'scale')),
            'FitToWidth' => self::TryParseInt($pageSetup === null ? null : self::AttributeValue($pageSetup, 'fitToWidth')),
            'FitToHeight' => self::TryParseInt($pageSetup === null ? null : self::AttributeValue($pageSetup, 'fitToHeight')),
            'PrintGridLines' => self::ParseBool($printOptions === null ? null : self::AttributeValue($printOptions, 'gridLines')),
            'PrintHeadings' => self::ParseBool($printOptions === null ? null : self::AttributeValue($printOptions, 'headings')),
            'HorizontalCentered' => self::ParseBool($printOptions === null ? null : self::AttributeValue($printOptions, 'horizontalCentered')),
            'VerticalCentered' => self::ParseBool($printOptions === null ? null : self::AttributeValue($printOptions, 'verticalCentered')),
            'OddHeader' => $headerFooter === null ? '' : (string)($headerFooter->Element(self::MainNs . 'oddHeader') === null ? '' : $headerFooter->Element(self::MainNs . 'oddHeader')->Value),
            'OddFooter' => $headerFooter === null ? '' : (string)($headerFooter->Element(self::MainNs . 'oddFooter') === null ? '' : $headerFooter->Element(self::MainNs . 'oddFooter')->Value),
            'RowBreaks' => $rowBreaks,
            'ColumnBreaks' => $columnBreaks,
        );
    }

    private static function CreateOpenXmlPageSetupWorkbook($path)
    {
        $worksheetXml = '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:C10"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Title</t></is></c></row><row r="10"><c r="C10"><v>42</v></c></row></sheetData><printOptions headings="1" gridLines="1" gridLinesSet="1" horizontalCentered="1" verticalCentered="1"/><pageMargins left="0.25" right="0.4" top="0.5" bottom="0.6" header="0.2" footer="0.22"/><pageSetup paperSize="9" scale="95" fitToWidth="1" fitToHeight="2" firstPageNumber="3" useFirstPageNumber="1" orientation="landscape"/><headerFooter><oddHeader>&amp;LLeft Header&amp;CCenter Header&amp;RRight Header</oddHeader><oddFooter>&amp;LLeft Footer&amp;CCenter Footer&amp;RRight Footer</oddFooter></headerFooter><rowBreaks count="2" manualBreakCount="2"><brk id="4" max="16383" man="1"/><brk id="7" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks></worksheet>';
        $workbookXml = '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Print Sheet" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">\'Print Sheet\'!$A$1:$C$10</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">\'Print Sheet\'!$1:$2,\'Print Sheet\'!$A:$B</definedName></definedNames></workbook>';

        $contentTypesXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';

        $rootRelsXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

        $workbookRelsXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';

        ZipPackageHelper::CreatePackage($path, array(
            '[Content_Types].xml' => $contentTypesXml,
            '_rels/.rels' => $rootRelsXml,
            'xl/workbook.xml' => $workbookXml,
            'xl/_rels/workbook.xml.rels' => $workbookRelsXml,
            'xl/worksheets/sheet1.xml' => $worksheetXml,
        ));
    }

    private static function AttributeValue($element, $name)
    {
        $attribute = $element->Attribute($name);
        return $attribute === null ? null : (string)$attribute->Value;
    }

    private static function AttributeIntOrZero($element, $name)
    {
        $attribute = $element->Attribute($name);
        return intval($attribute === null ? '0' : (string)$attribute->Value);
    }

    private static function UnqualifyDefinedNameValue($value)
    {
        $trimmed = trim($value);
        $exclamationIndex = strrpos($trimmed, '!');
        return $exclamationIndex >= 0 ? substr($trimmed, $exclamationIndex + 1) : $trimmed;
    }

    private static function TryParseInt($value)
    {
        if ($value === null || !is_numeric($value)) {
            return null;
        }

        return intval($value);
    }

    private static function TryParseDouble($value)
    {
        if ($value === null || !is_numeric($value)) {
            return null;
        }

        return floatval($value);
    }

    private static function ParseBool($value)
    {
        if ($value === null) {
            return false;
        }

        return $value === '1' || strcasecmp($value, 'true') === 0;
    }

    private static function ResolvePartUri($sourceUri, $target)
    {
        if (str_starts_with($target, '/')) {
            return $target;
        }

        $slashIndex = strrpos($sourceUri, '/');
        $directory = $slashIndex === false ? '/' : substr($sourceUri, 0, $slashIndex + 1);
        $path = $directory . $target;

        $segments = explode('/', $path);
        $resolved = array();
        foreach ($segments as $segment) {
            if ($segment === '' || $segment === '.') {
                continue;
            }

            if ($segment === '..') {
                if (count($resolved) > 0) {
                    array_pop($resolved);
                }
                continue;
            }

            $resolved[] = $segment;
        }

        return '/' . implode('/', $resolved);
    }
}

Program::Main();
