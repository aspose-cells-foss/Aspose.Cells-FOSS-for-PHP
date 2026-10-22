<?php

namespace Aspose\Cells_FOSS\ConditionalFormattingOpenXmlTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\FormatConditionType;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\ConditionalFormattingScenarioFactory;
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
            'ConditionalFormatting.OpenXmlTests',
            array(
                new TestCase('conditional_formattings_match_openxml_sdk', array(Program::class, 'ConditionalFormattingsMatchOpenXmlSdk')),
                new TestCase('advanced_conditional_formattings_match_openxml_sdk', array(Program::class, 'AdvancedConditionalFormattingsMatchOpenXmlSdk')),
                new TestCase('openxml_created_conditional_formattings_load_in_library', array(Program::class, 'OpenXmlCreatedConditionalFormattingsLoadInLibrary')),
            )
        );
    }

    public static function ConditionalFormattingsMatchOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('conditional-formatting-openxml-write');
        $path = $temp->GetPath('conditional-formatting.xlsx');

        $workbook = ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        ConditionalFormattingScenarioFactory::AssertConditionalFormattings($loaded);

        $snapshot = self::ReadSnapshots($path, 'Conditional Formatting');
        AssertEx::Equal(2, count($snapshot));
        self::AssertSnapshot(self::GetSnapshot($snapshot, 'A1:A5'), 'A1:A5', 2);
        self::AssertSnapshot(self::GetSnapshot($snapshot, 'C1:C3'), 'C1:C3', 1);
    }

    public static function AdvancedConditionalFormattingsMatchOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('conditional-formatting-openxml-advanced-write');
        $path = $temp->GetPath('advanced-conditional-formatting.xlsx');

        $workbook = ConditionalFormattingScenarioFactory::CreateAdvancedConditionalFormattingWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        ConditionalFormattingScenarioFactory::AssertAdvancedConditionalFormattings($loaded);

        $worksheetXml = ZipPackageHelper::ReadEntryText($path, 'xl/worksheets/sheet1.xml');
        AssertEx::Contains('type="containsText"', $worksheetXml);
        AssertEx::Contains('text="error"', $worksheetXml);
        AssertEx::Contains('type="timePeriod"', $worksheetXml);
        AssertEx::Contains('timePeriod="today"', $worksheetXml);
        AssertEx::Contains('type="duplicateValues"', $worksheetXml);
        AssertEx::Contains('type="uniqueValues"', $worksheetXml);
        AssertEx::Contains('type="colorScale"', $worksheetXml);
        AssertEx::Contains('type="dataBar"', $worksheetXml);
        AssertEx::Contains('type="iconSet"', $worksheetXml);
    }

    public static function OpenXmlCreatedConditionalFormattingsLoadInLibrary()
    {
        $temp = new TemporaryDirectory('conditional-formatting-openxml-read');
        $path = $temp->GetPath('openxml-conditional-formatting.xlsx');

        self::CreateOpenXmlWorkbook($path);

        $loaded = new Workbook($path);
        AssertEx::Equal(1, $loaded->Worksheets[0]->ConditionalFormattings->Count);
        $collection = $loaded->Worksheets[0]->ConditionalFormattings[0];
        AssertEx::Equal(2, $collection->Count);
        AssertEx::Equal(1, $collection->RangeCount);
        AssertEx::Equal(FormatConditionType::CellValue, $collection[0]->Type);
        AssertEx::Equal(OperatorType::GreaterThan, $collection[0]->Operator);
        AssertEx::Equal('10', $collection[0]->Formula1);
        AssertEx::Equal(1, $collection[0]->Priority);
        AssertEx::True($collection[0]->StopIfTrue);
        AssertEx::Equal(FillPattern::Solid, $collection[0]->Style->Pattern);
        AssertEx::Equal(Color::FromArgb(255, 255, 235, 156), $collection[0]->Style->ForegroundColor);
        AssertEx::Equal(FormatConditionType::Expression, $collection[1]->Type);
        AssertEx::Equal('A1="Y"', $collection[1]->Formula1);
        AssertEx::Equal(2, $collection[1]->Priority);
    }

    private static function ReadSnapshots($path, $sheetName)
    {
        $archive = new \ZipArchive();
        if ($archive->open($path) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $workbookXml = $archive->getFromName('xl/workbook.xml');
        $relationshipsXml = $archive->getFromName('xl/_rels/workbook.xml.rels');
        $archive->close();

        if ($workbookXml === false) {
            throw new \RuntimeException('The workbook part was not found.');
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
        foreach ($workbookDocument->Root->Elements(self::MainNs . 'sheets') as $sheetsElement) {
            foreach ($sheetsElement->Elements(self::MainNs . 'sheet') as $candidate) {
                if ((string)$candidate->Attribute('name') === $sheetName) {
                    $sheetElement = $candidate;
                    break 2;
                }
            }
        }

        if ($sheetElement === null) {
            throw new \RuntimeException('Worksheet was not found.');
        }

        $relationshipId = (string)$sheetElement->Attribute(self::RelationshipNs . 'id');
        if ($relationshipId === '') {
            throw new \RuntimeException('Worksheet relationship id is missing.');
        }

        if (!array_key_exists($relationshipId, $relationships)) {
            throw new \RuntimeException('Worksheet relationship was not found.');
        }

        $worksheetUri = ltrim($relationships[$relationshipId], '/');

        $archive = new \ZipArchive();
        if ($archive->open($path) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $worksheetXml = $archive->getFromName($worksheetUri);
        $archive->close();
        if ($worksheetXml === false) {
            throw new \RuntimeException('Worksheet part was not found.');
        }

        $worksheetDocument = XDocument::Load($worksheetXml);

        $snapshots = array();
        foreach ($worksheetDocument->Root->Elements(self::MainNs . 'conditionalFormatting') as $formatting) {
            $rules = array();
            foreach ($formatting->Elements(self::MainNs . 'cfRule') as $rule) {
                $formulas = array();
                foreach ($rule->Elements(self::MainNs . 'formula') as $formula) {
                    $formulas[] = $formula->Value;
                }

                $rules[] = array(
                    'Type' => self::AttributeOrEmpty($rule, 'type'),
                    'Operator' => self::AttributeOrEmpty($rule, 'operator'),
                    'DxfId' => self::AttributeOrEmpty($rule, 'dxfId'),
                    'Priority' => self::AttributeOrEmpty($rule, 'priority'),
                    'StopIfTrue' => self::AttributeOrEmpty($rule, 'stopIfTrue'),
                    'Formulas' => $formulas,
                );
            }

            $snapshots[] = array(
                'Sqref' => strtoupper(self::AttributeOrEmpty($formatting, 'sqref')),
                'Rules' => $rules,
            );
        }

        return $snapshots;
    }

    private static function AttributeOrEmpty($element, $name)
    {
        $attribute = $element->Attribute($name);
        return $attribute === null ? '' : (string)$attribute->Value;
    }

    private static function GetSnapshot($snapshots, $sqref)
    {
        foreach ($snapshots as $snapshot) {
            if ($snapshot['Sqref'] === $sqref) {
                return $snapshot;
            }
        }

        throw new \RuntimeException('Snapshot was not found: ' . $sqref);
    }

    private static function AssertSnapshot($snapshot, $sqref, $ruleCount)
    {
        AssertEx::Equal($sqref, $snapshot['Sqref']);
        AssertEx::Equal($ruleCount, count($snapshot['Rules']));
    }

    private static function CreateOpenXmlWorkbook($path)
    {
        $worksheetXml = '<?xml version="1.0" encoding="utf-8"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<dimension ref="A1:A3"/>'
            . '<sheetData>'
            . '<row r="1"><c r="A1" t="inlineStr"><is><t>Y</t></is></c></row>'
            . '</sheetData>'
            . '<conditionalFormatting sqref="A1:A3">'
            . '<cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1" stopIfTrue="1"><formula>10</formula></cfRule>'
            . '<cfRule type="expression" priority="2"><formula>A1="Y"</formula></cfRule>'
            . '</conditionalFormatting>'
            . '</worksheet>';

        $workbookXml = '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Conditional Formatting" sheetId="1" r:id="rId1"/></sheets></workbook>';

        $stylesXml = '<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill></dxf></dxfs></styleSheet>';

        $contentTypesXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';

        $rootRelsXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

        $workbookRelsXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

        ZipPackageHelper::CreatePackage($path, array(
            '[Content_Types].xml' => $contentTypesXml,
            '_rels/.rels' => $rootRelsXml,
            'xl/workbook.xml' => $workbookXml,
            'xl/_rels/workbook.xml.rels' => $workbookRelsXml,
            'xl/worksheets/sheet1.xml' => $worksheetXml,
            'xl/styles.xml' => $stylesXml,
        ));
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
