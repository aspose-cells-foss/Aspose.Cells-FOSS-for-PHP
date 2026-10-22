<?php

namespace Aspose\Cells_FOSS\DataValidationOpenXmlTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';

use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\OperatorType;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\ValidationScenarioFactory;
use Aspose\Cells_FOSS\Testing\ZipPackageHelper;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    public static function Main()
    {
        return TestRunner::Run(
            'DataValidation.OpenXmlTests',
            array(
                new TestCase('data_validations_match_openxml_sdk', array(Program::class, 'DataValidationsMatchOpenXmlSdk')),
                new TestCase('openxml_created_data_validations_load_in_library', array(Program::class, 'OpenXmlCreatedDataValidationsLoadInLibrary')),
            )
        );
    }

    public static function DataValidationsMatchOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('validation-openxml-write');
        $path = $temp->GetPath('validations.xlsx');

        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        ValidationScenarioFactory::AssertValidations($loaded);

        $snapshot = self::ReadValidationSnapshots($path, 'Validation Sheet');
        AssertEx::Equal(3, count($snapshot));

        self::AssertValidation(self::GetValidationSnapshot($snapshot, 'A1:A3'), 'A1:A3', 'list', '', 'stop', true, false, true, true, 'Status', 'Pick a status', 'Invalid', 'Choose from the list', '"Open,Closed"', '');
        self::AssertValidation(self::GetValidationSnapshot($snapshot, 'B2:C3 E2:E3'), 'B2:C3 E2:E3', 'decimal', 'between', 'stop', false, true, false, true, '', '', 'Range', 'Enter 1.5-9.5', '1.5', '9.5');
        self::AssertValidation(self::GetValidationSnapshot($snapshot, 'G1'), 'G1', 'custom', '', 'warning', false, false, true, false, 'Code', 'Up to 5 chars', '', '', 'LEN(G1)<=5', '');
    }

    public static function OpenXmlCreatedDataValidationsLoadInLibrary()
    {
        $temp = new TemporaryDirectory('validation-openxml-read');
        $path = $temp->GetPath('openxml-validations.xlsx');

        self::CreateOpenXmlValidationWorkbook($path);

        $loaded = new Workbook($path);
        $sheet = $loaded->Worksheets[0];
        AssertEx::Equal('Validation Sheet', $sheet->Name);
        AssertEx::Equal(2, $sheet->Validations->Count);

        $listValidation = $sheet->Validations[0];
        AssertEx::Equal(ValidationType::List, $listValidation->Type);
        AssertEx::Equal('"One,Two"', $listValidation->Formula1);
        AssertEx::True($listValidation->IgnoreBlank);
        AssertEx::True($listValidation->ShowInput);
        AssertEx::Equal('Status', $listValidation->InputTitle);
        AssertEx::Equal('Pick one', $listValidation->InputMessage);
        AssertEx::Equal(1, count($listValidation->Areas));
        AssertEx::Equal(0, $listValidation->Areas[0]->StartRow);
        AssertEx::Equal(0, $listValidation->Areas[0]->StartColumn);
        AssertEx::Equal(1, $listValidation->Areas[0]->EndRow);
        AssertEx::Equal(0, $listValidation->Areas[0]->EndColumn);

        $decimalValidation = $sheet->Validations[1];
        AssertEx::Equal(ValidationType::Decimal, $decimalValidation->Type);
        AssertEx::Equal(OperatorType::Between, $decimalValidation->Operator);
        AssertEx::Equal('1', $decimalValidation->Formula1);
        AssertEx::Equal('9', $decimalValidation->Formula2);
        AssertEx::False($decimalValidation->InCellDropDown);
        AssertEx::True($decimalValidation->ShowError);
        AssertEx::Equal('Range', $decimalValidation->ErrorTitle);
        AssertEx::Equal('Enter 1-9', $decimalValidation->ErrorMessage);
        AssertEx::Equal(2, count($decimalValidation->Areas));
        AssertEx::Equal(1, $decimalValidation->Areas[0]->StartRow);
        AssertEx::Equal(1, $decimalValidation->Areas[0]->StartColumn);
        AssertEx::Equal(2, $decimalValidation->Areas[0]->EndRow);
        AssertEx::Equal(2, $decimalValidation->Areas[0]->EndColumn);
        AssertEx::Equal(1, $decimalValidation->Areas[1]->StartRow);
        AssertEx::Equal(4, $decimalValidation->Areas[1]->StartColumn);
        AssertEx::Equal(2, $decimalValidation->Areas[1]->EndRow);
        AssertEx::Equal(4, $decimalValidation->Areas[1]->EndColumn);

        $snapshot = self::ReadValidationSnapshots($path, 'Validation Sheet');
        AssertEx::Equal(2, count($snapshot));
        self::AssertValidation(self::GetValidationSnapshot($snapshot, 'A1:A2'), 'A1:A2', 'list', '', 'stop', true, false, true, false, 'Status', 'Pick one', '', '', '"One,Two"', '');
        self::AssertValidation(self::GetValidationSnapshot($snapshot, 'B2:C3 E2:E3'), 'B2:C3 E2:E3', 'decimal', 'between', 'stop', false, true, false, true, '', '', 'Range', 'Enter 1-9', '1', '9');
    }

    private static function ReadValidationSnapshots($workbookPath, $sheetName)
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
        $validations = array();

        $validationsElement = $worksheetDocument->Root === null ? null : $worksheetDocument->Root->Element(self::MainNs . 'dataValidations');
        $validationElements = $validationsElement === null ? array() : $validationsElement->Elements(self::MainNs . 'dataValidation');
        foreach ($validationElements as $validation) {
            $sqrefAttribute = $validation->Attribute('sqref');
            $typeAttribute = $validation->Attribute('type');
            $operatorAttribute = $validation->Attribute('operator');
            $errorStyleAttribute = $validation->Attribute('errorStyle');
            $allowBlankAttribute = $validation->Attribute('allowBlank');
            $showDropDownAttribute = $validation->Attribute('showDropDown');
            $showInputMessageAttribute = $validation->Attribute('showInputMessage');
            $showErrorMessageAttribute = $validation->Attribute('showErrorMessage');
            $promptTitleAttribute = $validation->Attribute('promptTitle');
            $promptAttribute = $validation->Attribute('prompt');
            $errorTitleAttribute = $validation->Attribute('errorTitle');
            $errorAttribute = $validation->Attribute('error');

            $validations[] = array(
                'Sqref' => strtoupper($sqrefAttribute === null ? '' : (string)$sqrefAttribute->Value),
                'Type' => $typeAttribute === null ? '' : (string)$typeAttribute->Value,
                'Operator' => $operatorAttribute === null ? '' : (string)$operatorAttribute->Value,
                'ErrorStyle' => $errorStyleAttribute === null ? 'stop' : (string)$errorStyleAttribute->Value,
                'AllowBlank' => self::ParseBool($allowBlankAttribute === null ? null : (string)$allowBlankAttribute->Value),
                'ShowDropDown' => self::ParseBool($showDropDownAttribute === null ? null : (string)$showDropDownAttribute->Value),
                'ShowInputMessage' => self::ParseBool($showInputMessageAttribute === null ? null : (string)$showInputMessageAttribute->Value),
                'ShowErrorMessage' => self::ParseBool($showErrorMessageAttribute === null ? null : (string)$showErrorMessageAttribute->Value),
                'PromptTitle' => $promptTitleAttribute === null ? '' : (string)$promptTitleAttribute->Value,
                'Prompt' => $promptAttribute === null ? '' : (string)$promptAttribute->Value,
                'ErrorTitle' => $errorTitleAttribute === null ? '' : (string)$errorTitleAttribute->Value,
                'Error' => $errorAttribute === null ? '' : (string)$errorAttribute->Value,
                'Formula1' => self::ElementValueOrEmpty($validation->Element(self::MainNs . 'formula1')),
                'Formula2' => self::ElementValueOrEmpty($validation->Element(self::MainNs . 'formula2')),
            );
        }

        return $validations;
    }

    private static function CreateOpenXmlValidationWorkbook($path)
    {
        $worksheetXml = '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>One</t></is></c></row></sheetData><dataValidations count="2"><dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Status" prompt="Pick one" sqref="A1:A2"><formula1>"One,Two"</formula1></dataValidation><dataValidation type="decimal" operator="between" showDropDown="1" showErrorMessage="1" errorTitle="Range" error="Enter 1-9" sqref="B2:C3 E2:E3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations></worksheet>';
        $workbookXml = '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Validation Sheet" sheetId="1" r:id="rId1"/></sheets></workbook>';

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

    private static function ElementValueOrEmpty($element)
    {
        return $element === null ? '' : (string)$element->Value;
    }

    private static function GetValidationSnapshot($snapshots, $sqref)
    {
        foreach ($snapshots as $snapshot) {
            if ($snapshot['Sqref'] === $sqref) {
                return $snapshot;
            }
        }

        throw new \RuntimeException('Snapshot was not found: ' . $sqref);
    }

    private static function AssertValidation($actual, $sqref, $type, $operatorName, $errorStyle, $allowBlank, $showDropDown, $showInputMessage, $showErrorMessage, $promptTitle, $prompt, $errorTitle, $error, $formula1, $formula2)
    {
        AssertEx::Equal($sqref, $actual['Sqref']);
        AssertEx::Equal($type, $actual['Type']);
        AssertEx::Equal($operatorName, $actual['Operator']);
        AssertEx::Equal($errorStyle, $actual['ErrorStyle']);
        AssertEx::Equal($allowBlank, $actual['AllowBlank']);
        AssertEx::Equal($showDropDown, $actual['ShowDropDown']);
        AssertEx::Equal($showInputMessage, $actual['ShowInputMessage']);
        AssertEx::Equal($showErrorMessage, $actual['ShowErrorMessage']);
        AssertEx::Equal($promptTitle, $actual['PromptTitle']);
        AssertEx::Equal($prompt, $actual['Prompt']);
        AssertEx::Equal($errorTitle, $actual['ErrorTitle']);
        AssertEx::Equal($error, $actual['Error']);
        AssertEx::Equal($formula1, $actual['Formula1']);
        AssertEx::Equal($formula2, $actual['Formula2']);
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
