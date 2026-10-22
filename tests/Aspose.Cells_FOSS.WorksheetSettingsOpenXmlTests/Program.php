<?php

namespace Aspose\Cells_FOSS\WorksheetSettingsOpenXmlTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';
require_once __DIR__ . '/../Shared/AutoFilterScenarioFactory.php';

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\AutoFilterScenarioFactory;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\WorksheetScenarioFactory;
use Aspose\Cells_FOSS\Testing\ZipPackageHelper;
use Aspose\Cells_FOSS\VisibilityType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const RelationshipNs = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    public static function Main()
    {
        return TestRunner::Run(
            'WorksheetSettings.OpenXmlTests',
            array(
                new TestCase('worksheet_visibility_matches_openxml_sdk', array(Program::class, 'WorksheetVisibilityMatchesOpenXmlSdk')),
                new TestCase('worksheet_settings_and_values_match_openxml_sdk', array(Program::class, 'WorksheetSettingsAndValuesMatchOpenXmlSdk')),
                new TestCase('worksheet_view_and_tab_color_match_openxml_sdk', array(Program::class, 'WorksheetViewAndTabColorMatchOpenXmlSdk')),
                new TestCase('worksheet_protection_match_openxml_sdk', array(Program::class, 'WorksheetProtectionMatchOpenXmlSdk')),
                new TestCase('autofilter_matches_openxml_sdk', array(Program::class, 'AutoFilterMatchesOpenXmlSdk')),
                new TestCase('adjacent_columns_are_combined_and_read_consistently', array(Program::class, 'AdjacentColumnsAreCombinedAndReadConsistently')),
                new TestCase('merge_only_dimension_matches_openxml_sdk', array(Program::class, 'MergeOnlyDimensionMatchesOpenXmlSdk')),
                new TestCase('openxml_created_visibility_loads_in_library', array(Program::class, 'OpenXmlCreatedVisibilityLoadsInLibrary')),
                new TestCase('openxml_created_settings_and_values_load_in_library', array(Program::class, 'OpenXmlCreatedSettingsAndValuesLoadInLibrary')),
                new TestCase('openxml_created_sheet_view_loads_in_library', array(Program::class, 'OpenXmlCreatedSheetViewLoadsInLibrary')),
                new TestCase('openxml_created_sheet_protection_loads_in_library', array(Program::class, 'OpenXmlCreatedSheetProtectionLoadsInLibrary')),
                new TestCase('openxml_created_autofilter_loads_in_library', array(Program::class, 'OpenXmlCreatedAutoFilterLoadsInLibrary')),
                new TestCase('openxml_created_columns_load_in_library', array(Program::class, 'OpenXmlCreatedColumnsLoadInLibrary')),
                new TestCase('openxml_created_merge_only_dimension_loads_in_library', array(Program::class, 'OpenXmlCreatedMergeOnlyDimensionLoadsInLibrary')),
            )
        );
    }

    public static function WorksheetVisibilityMatchesOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-visibility');
        $path = $temp->GetPath('visibility.xlsx');

        $workbook = new Workbook();
        $workbook->Worksheets[0]->Name = 'Visible';
        $hiddenIndex = $workbook->Worksheets->Add('Hidden');
        $veryHiddenIndex = $workbook->Worksheets->Add('VeryHidden');
        $workbook->Worksheets[$hiddenIndex]->VisibilityType = VisibilityType::Hidden;
        $workbook->Worksheets[$veryHiddenIndex]->VisibilityType = VisibilityType::VeryHidden;
        $workbook->Save($path);

        $loaded = new Workbook($path);
        AssertEx::Equal(VisibilityType::Visible, $loaded->Worksheets['Visible']->VisibilityType);
        AssertEx::Equal(VisibilityType::Hidden, $loaded->Worksheets['Hidden']->VisibilityType);
        AssertEx::Equal(VisibilityType::VeryHidden, $loaded->Worksheets['VeryHidden']->VisibilityType);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(3, count($snapshots));
        AssertEx::Equal('Visible', $snapshots[0]['Name']);
        AssertEx::Equal('Visible', $snapshots[0]['Visibility']);
        AssertEx::Equal('Hidden', $snapshots[1]['Name']);
        AssertEx::Equal('Hidden', $snapshots[1]['Visibility']);
        AssertEx::Equal('VeryHidden', $snapshots[2]['Name']);
        AssertEx::Equal('VeryHidden', $snapshots[2]['Visibility']);
    }

    public static function WorksheetSettingsAndValuesMatchOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-settings');
        $path = $temp->GetPath('worksheet-settings.xlsx');

        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(2, count($snapshots));
        $snapshot = self::GetSnapshotByName($snapshots, 'Layout');
        AssertEx::Equal('Layout', $snapshot['Name']);
        AssertEx::Equal('Hidden', $snapshot['Visibility']);
        AssertEx::Equal('A1:C4', $snapshot['Dimension']);
        AssertEx::Equal('Merged', $snapshot['Cells']['A1']);
        AssertEx::Equal('99', $snapshot['Cells']['C4']);
        self::AssertRow($snapshot, 1, 22.5, false);
        self::AssertRow($snapshot, 3, null, true);
        self::AssertColumn($snapshot, 0, 0, 18.25, false);
        self::AssertColumn($snapshot, 2, 2, null, true);
        AssertEx::Equal(1, count($snapshot['MergeReferences']));
        AssertEx::Equal('A1:B2', $snapshot['MergeReferences'][0]);

        $layoutSheet = $loaded->Worksheets['Layout'];
        AssertEx::NotNull($layoutSheet);
        AssertEx::Equal($layoutSheet->Cells['A1']->StringValue, $snapshot['Cells']['A1']);
        AssertEx::Equal($layoutSheet->Cells['C4']->StringValue, $snapshot['Cells']['C4']);
        AssertEx::Equal($layoutSheet->Cells->Rows[1]->Height ?? 0.0, $snapshot['Rows'][1]['Height'] ?? 0.0);
        AssertEx::Equal($layoutSheet->Cells->Rows[3]->IsHidden, $snapshot['Rows'][3]['Hidden']);
        AssertEx::Equal($layoutSheet->Cells->Columns[0]->Width ?? 0.0, $snapshot['Columns'][0]['Width'] ?? 0.0);
        AssertEx::Equal($layoutSheet->Cells->Columns[2]->IsHidden, $snapshot['Columns'][1]['Hidden']);
    }

    public static function WorksheetViewAndTabColorMatchOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-view');
        $path = $temp->GetPath('worksheet-view.xlsx');

        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(2, count($snapshots));
        $snapshot = self::GetSnapshotByName($snapshots, 'Layout');
        AssertEx::Equal('FF224466', $snapshot['TabColor']);
        AssertEx::False($snapshot['ShowGridLines']);
        AssertEx::False($snapshot['ShowRowColumnHeaders']);
        AssertEx::False($snapshot['ShowZeros']);
        AssertEx::True($snapshot['RightToLeft']);
        AssertEx::Equal(85, $snapshot['Zoom']);
    }

    public static function WorksheetProtectionMatchOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-protection');
        $path = $temp->GetPath('worksheet-protection.xlsx');

        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(2, count($snapshots));
        $snapshot = self::GetSnapshotByName($snapshots, 'Layout');
        AssertEx::True($snapshot['IsProtected']);
        AssertEx::True($snapshot['ProtectObjects']);
        AssertEx::True($snapshot['ProtectScenarios']);
        AssertEx::True($snapshot['ProtectFormatCells']);
        AssertEx::True($snapshot['ProtectInsertRows']);
        AssertEx::True($snapshot['ProtectAutoFilter']);
        AssertEx::True($snapshot['ProtectSelectLockedCells']);
        AssertEx::True($snapshot['ProtectSelectUnlockedCells']);
    }

    public static function AutoFilterMatchesOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-autofilter');
        $path = $temp->GetPath('autofilter.xlsx');

        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        $workbook->Save($path);

        $loaded = new Workbook($path);
        AutoFilterScenarioFactory::AssertAutoFilter($loaded);

        $archive = new \ZipArchive();
        if ($archive->open($path) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $worksheetXml = $archive->getFromName('xl/worksheets/sheet1.xml');
        $archive->close();
        if ($worksheetXml === false) {
            throw new \RuntimeException('Worksheet part was not found.');
        }

        $worksheetDocument = XDocument::Load($worksheetXml);
        $autoFilter = $worksheetDocument->Root === null ? null : $worksheetDocument->Root->Element(self::MainNs . 'autoFilter');
        AssertEx::NotNull($autoFilter);
        $refAttribute = $autoFilter->Attribute('ref');
        AssertEx::Equal('A1:E6', strtoupper($refAttribute === null ? '' : (string)$refAttribute->Value));

        $filterColumns = $autoFilter->Elements(self::MainNs . 'filterColumn');
        AssertEx::Equal(5, count($filterColumns));

        $hiddenButtonAttribute = $filterColumns[0]->Attribute('hiddenButton');
        AssertEx::Equal('1', $hiddenButtonAttribute === null ? '' : (string)$hiddenButtonAttribute->Value);
        $filters = $filterColumns[0]->Element(self::MainNs . 'filters');
        $firstFilter = $filters === null ? null : $filters->Element(self::MainNs . 'filter');
        $firstFilterVal = $firstFilter === null ? null : $firstFilter->Attribute('val');
        AssertEx::Equal('Open', $firstFilterVal === null ? '' : (string)$firstFilterVal->Value);

        $customFilters = $filterColumns[1]->Element(self::MainNs . 'customFilters');
        $firstCustomFilter = $customFilters === null ? null : $customFilters->Element(self::MainNs . 'customFilter');
        $firstCustomOperator = $firstCustomFilter === null ? null : $firstCustomFilter->Attribute('operator');
        AssertEx::Equal('greaterThanOrEqual', $firstCustomOperator === null ? '' : (string)$firstCustomOperator->Value);

        $colorFilter = $filterColumns[2]->Element(self::MainNs . 'colorFilter');
        $colorFilterDxfId = $colorFilter === null ? null : $colorFilter->Attribute('dxfId');
        AssertEx::Equal('3', $colorFilterDxfId === null ? '' : (string)$colorFilterDxfId->Value);

        $dynamicFilter = $filterColumns[3]->Element(self::MainNs . 'dynamicFilter');
        $dynamicFilterType = $dynamicFilter === null ? null : $dynamicFilter->Attribute('type');
        AssertEx::Equal('thisMonth', $dynamicFilterType === null ? '' : (string)$dynamicFilterType->Value);

        $top10 = $filterColumns[4]->Element(self::MainNs . 'top10');
        $top10Top = $top10 === null ? null : $top10->Attribute('top');
        AssertEx::Equal('0', $top10Top === null ? '' : (string)$top10Top->Value);

        $sortState = $autoFilter->Element(self::MainNs . 'sortState');
        $sortStateRef = $sortState === null ? null : $sortState->Attribute('ref');
        AssertEx::Equal('A2:E6', strtoupper($sortStateRef === null ? '' : (string)$sortStateRef->Value));

        $sortConditions = $sortState === null ? array() : $sortState->Elements(self::MainNs . 'sortCondition');
        AssertEx::Equal(3, count($sortConditions));

        $sortByAttribute = $sortConditions[0]->Attribute('sortBy');
        AssertEx::Equal('value', $sortByAttribute === null ? '' : (string)$sortByAttribute->Value);
        $customListAttribute = $sortConditions[0]->Attribute('customList');
        AssertEx::Equal('High,Medium,Low', $customListAttribute === null ? '' : (string)$customListAttribute->Value);

        $sortByAttribute = $sortConditions[1]->Attribute('sortBy');
        AssertEx::Equal('cellColor', $sortByAttribute === null ? '' : (string)$sortByAttribute->Value);
        $dxfIdAttribute = $sortConditions[1]->Attribute('dxfId');
        AssertEx::Equal('4', $dxfIdAttribute === null ? '' : (string)$dxfIdAttribute->Value);

        $sortByAttribute = $sortConditions[2]->Attribute('sortBy');
        AssertEx::Equal('icon', $sortByAttribute === null ? '' : (string)$sortByAttribute->Value);
        $iconSetAttribute = $sortConditions[2]->Attribute('iconSet');
        AssertEx::Equal('3TrafficLights1', $iconSetAttribute === null ? '' : (string)$iconSetAttribute->Value);
        $iconIdAttribute = $sortConditions[2]->Attribute('iconId');
        AssertEx::Equal('2', $iconIdAttribute === null ? '' : (string)$iconIdAttribute->Value);
    }

    public static function AdjacentColumnsAreCombinedAndReadConsistently()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-columns');
        $path = $temp->GetPath('columns.xlsx');

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Columns';
        $sheet->Cells->Columns[0]->Width = 15.5;
        $sheet->Cells->Columns[1]->Width = 15.5;
        $sheet->Cells->Columns[3]->IsHidden = true;
        $sheet->Cells['A1']->PutValue('A');
        $sheet->Cells['B1']->PutValue('B');
        $workbook->Save($path);

        $loaded = new Workbook($path);
        AssertEx::Equal(15.5, $loaded->Worksheets[0]->Cells->Columns[0]->Width ?? 0.0);
        AssertEx::Equal(15.5, $loaded->Worksheets[0]->Cells->Columns[1]->Width ?? 0.0);
        AssertEx::True($loaded->Worksheets[0]->Cells->Columns[3]->IsHidden);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(1, count($snapshots));
        $snapshot = $snapshots[0];
        AssertEx::Equal(2, count($snapshot['Columns']));
        self::AssertColumn($snapshot, 0, 1, 15.5, false);
        self::AssertColumn($snapshot, 3, 3, null, true);
        AssertEx::Equal('A', $snapshot['Cells']['A1']);
        AssertEx::Equal('B', $snapshot['Cells']['B1']);
    }

    public static function MergeOnlyDimensionMatchesOpenXmlSdk()
    {
        $temp = new TemporaryDirectory('worksheet-openxml-merge-only');
        $path = $temp->GetPath('merge-only.xlsx');

        $workbook = new Workbook();
        $workbook->Worksheets[0]->Name = 'MergeOnly';
        $workbook->Worksheets[0]->Cells->Merge(4, 5, 2, 2);
        $workbook->Save($path);

        $loaded = new Workbook($path);
        AssertEx::Equal(1, count($loaded->Worksheets[0]->Cells->MergedCells));
        AssertEx::Equal(4, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartRow);
        AssertEx::Equal(5, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartColumn);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(1, count($snapshots));
        $snapshot = $snapshots[0];
        AssertEx::Equal('F5:G6', $snapshot['Dimension']);
        AssertEx::Equal(1, count($snapshot['MergeReferences']));
        AssertEx::Equal('F5:G6', $snapshot['MergeReferences'][0]);
    }

    public static function OpenXmlCreatedVisibilityLoadsInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-visibility');
        $path = $temp->GetPath('openxml-visibility.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array('Name' => 'Visible', 'State' => null, 'WorksheetXml' => self::MinimalWorksheetXml()),
            array('Name' => 'Hidden', 'State' => 'hidden', 'WorksheetXml' => self::MinimalWorksheetXml()),
            array('Name' => 'VeryHidden', 'State' => 'veryHidden', 'WorksheetXml' => self::MinimalWorksheetXml()),
        ));

        $loaded = new Workbook($path);
        AssertEx::Equal(VisibilityType::Visible, $loaded->Worksheets['Visible']->VisibilityType);
        AssertEx::Equal(VisibilityType::Hidden, $loaded->Worksheets['Hidden']->VisibilityType);
        AssertEx::Equal(VisibilityType::VeryHidden, $loaded->Worksheets['VeryHidden']->VisibilityType);

        $snapshot = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(3, count($snapshot));
        AssertEx::Equal('Hidden', $snapshot[1]['Visibility']);
        AssertEx::Equal('VeryHidden', $snapshot[2]['Visibility']);
    }

    public static function OpenXmlCreatedSettingsAndValuesLoadInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-settings');
        $path = $temp->GetPath('openxml-settings.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array(
                'Name' => 'Layout',
                'State' => 'hidden',
                'WorksheetXml' => '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF224466"/></sheetPr><dimension ref="A1:C4"/><sheetViews><sheetView workbookViewId="0" showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1" zoomScale="85"/></sheetViews><sheetProtection sheet="1" objects="1" scenarios="1" formatCells="1" insertRows="1" autoFilter="1" selectLockedCells="1" selectUnlockedCells="1"/><cols><col min="1" max="1" width="18.25" customWidth="1"/><col min="3" max="3" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row><row r="2" ht="22.5" customHeight="1"/><row r="4" hidden="1"><c r="C4"><v>99</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>',
            ),
        ));

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(1, count($snapshots));
        $snapshot = $snapshots[0];
        AssertEx::Equal('Merged', $snapshot['Cells']['A1']);
        AssertEx::Equal('99', $snapshot['Cells']['C4']);
    }

    public static function OpenXmlCreatedSheetViewLoadsInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-view');
        $path = $temp->GetPath('openxml-view.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array(
                'Name' => 'ViewSheet',
                'State' => null,
                'WorksheetXml' => '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF0A141E"/></sheetPr><sheetViews><sheetView workbookViewId="0" showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1" zoomScale="140"/></sheetViews><sheetData/></worksheet>',
            ),
        ));

        $loaded = new Workbook($path);
        $sheet = $loaded->Worksheets[0];
        AssertEx::Equal(Color::FromArgb(255, 10, 20, 30), $sheet->TabColor);
        AssertEx::False($sheet->ShowGridlines);
        AssertEx::False($sheet->ShowRowColumnHeaders);
        AssertEx::False($sheet->ShowZeros);
        AssertEx::True($sheet->RightToLeft);
        AssertEx::Equal(140, $sheet->Zoom);
    }

    public static function OpenXmlCreatedSheetProtectionLoadsInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-protection');
        $path = $temp->GetPath('openxml-protection.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array(
                'Name' => 'ProtectedSheet',
                'State' => null,
                'WorksheetXml' => '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetProtection sheet="1" objects="1" scenarios="1" formatCells="1" insertRows="1" autoFilter="1" selectLockedCells="1" selectUnlockedCells="1"/><sheetData/></worksheet>',
            ),
        ));

        $loaded = new Workbook($path);
        $sheet = $loaded->Worksheets[0];
        AssertEx::True($sheet->Protection->IsProtected);
        AssertEx::True($sheet->Protection->Objects);
        AssertEx::True($sheet->Protection->Scenarios);
        AssertEx::True($sheet->Protection->FormatCells);
        AssertEx::True($sheet->Protection->InsertRows);
        AssertEx::True($sheet->Protection->AutoFilter);
        AssertEx::True($sheet->Protection->SelectLockedCells);
        AssertEx::True($sheet->Protection->SelectUnlockedCells);
    }

    public static function OpenXmlCreatedAutoFilterLoadsInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-autofilter');
        $path = $temp->GetPath('openxml-autofilter.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array(
                'Name' => 'Filtered',
                'State' => null,
                'WorksheetXml' => '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E6"/><sheetData/><autoFilter ref="A1:E6"><filterColumn colId="0" hiddenButton="1"><filters><filter val="Open"/><filter val="Closed"/></filters></filterColumn><filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="10"/><customFilter operator="lessThanOrEqual" val="50"/></customFilters></filterColumn><filterColumn colId="2"><colorFilter dxfId="3" cellColor="1"/></filterColumn><filterColumn colId="3"><dynamicFilter type="thisMonth" val="1" maxVal="31"/></filterColumn><filterColumn colId="4"><top10 top="0" percent="1" val="10" filterVal="2.5"/></filterColumn><sortState ref="A2:E6" caseSensitive="1" sortMethod="pinYin"><sortCondition ref="B2:B6" descending="1" sortBy="value" customList="High,Medium,Low"/><sortCondition ref="C2:C6" sortBy="cellColor" dxfId="4"/><sortCondition ref="E2:E6" sortBy="icon" iconSet="3TrafficLights1" iconId="2"/></sortState></autoFilter></worksheet>',
            ),
        ));
        self::AddAutoFilterStylesPart($path);

        $loaded = new Workbook($path);
        AutoFilterScenarioFactory::AssertAutoFilter($loaded);
    }

    private static function AddAutoFilterStylesPart($path)
    {
        $stylesXml = '<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="5"><dxf/><dxf/><dxf/><dxf/><dxf/></dxfs></styleSheet>';

        $archive = new \ZipArchive();
        if ($archive->open($path) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $relationships = $archive->getFromName('xl/_rels/workbook.xml.rels');
        $contentTypes = $archive->getFromName('[Content_Types].xml');
        $archive->addFromString('xl/styles.xml', $stylesXml);
        $archive->addFromString('xl/_rels/workbook.xml.rels', str_replace(
            '</Relationships>',
            '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
            $relationships
        ));
        $archive->addFromString('[Content_Types].xml', str_replace(
            '</Types>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
            $contentTypes
        ));
        $archive->close();
    }

    public static function OpenXmlCreatedColumnsLoadInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-columns');
        $path = $temp->GetPath('openxml-columns.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array(
                'Name' => 'Columns',
                'State' => null,
                'WorksheetXml' => '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><cols><col min="1" max="2" width="15.5" customWidth="1"/><col min="4" max="4" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r="B1" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>',
            ),
        ));

        $loaded = new Workbook($path);
        AssertEx::Equal(15.5, $loaded->Worksheets[0]->Cells->Columns[0]->Width ?? 0.0);
        AssertEx::Equal(15.5, $loaded->Worksheets[0]->Cells->Columns[1]->Width ?? 0.0);
        AssertEx::True($loaded->Worksheets[0]->Cells->Columns[3]->IsHidden);
        AssertEx::Equal('A', $loaded->Worksheets[0]->Cells['A1']->StringValue);
        AssertEx::Equal('B', $loaded->Worksheets[0]->Cells['B1']->StringValue);
    }

    public static function OpenXmlCreatedMergeOnlyDimensionLoadsInLibrary()
    {
        $temp = new TemporaryDirectory('worksheet-library-read-merge-only');
        $path = $temp->GetPath('openxml-merge-only.xlsx');

        self::CreateOpenXmlWorkbook($path, array(
            array(
                'Name' => 'MergeOnly',
                'State' => null,
                'WorksheetXml' => '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="F5:G6"/><sheetData/><mergeCells count="1"><mergeCell ref="F5:G6"/></mergeCells></worksheet>',
            ),
        ));

        $loaded = new Workbook($path);
        AssertEx::Equal(1, count($loaded->Worksheets[0]->Cells->MergedCells));
        AssertEx::Equal(4, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartRow);
        AssertEx::Equal(5, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartColumn);

        $snapshots = self::ReadWorkbookSnapshot($path);
        AssertEx::Equal(1, count($snapshots));
        $snapshot = $snapshots[0];
        AssertEx::Equal('F5:G6', $snapshot['Dimension']);
        AssertEx::Equal(1, count($snapshot['MergeReferences']));
        AssertEx::Equal('F5:G6', $snapshot['MergeReferences'][0]);
    }

    private static function AssertRow($snapshot, $rowIndex, $height, $hidden)
    {
        if (!array_key_exists($rowIndex, $snapshot['Rows'])) {
            throw new \RuntimeException('Expected row ' . $rowIndex . ' to exist in Open XML snapshot.');
        }

        $row = $snapshot['Rows'][$rowIndex];
        AssertEx::Equal($height, $row['Height']);
        AssertEx::Equal($hidden, $row['Hidden']);
    }

    private static function AssertColumn($snapshot, $minColumnIndex, $maxColumnIndex, $width, $hidden)
    {
        $column = null;
        foreach ($snapshot['Columns'] as $candidate) {
            if ($candidate['MinColumnIndex'] === $minColumnIndex && $candidate['MaxColumnIndex'] === $maxColumnIndex) {
                $column = $candidate;
                break;
            }
        }

        AssertEx::NotNull($column, 'Expected column range ' . $minColumnIndex . ':' . $maxColumnIndex . ' to exist in Open XML snapshot.');
        AssertEx::Equal($width, $column['Width']);
        AssertEx::Equal($hidden, $column['Hidden']);
    }

    private static function GetSnapshotByName($snapshots, $name)
    {
        foreach ($snapshots as $snapshot) {
            if ($snapshot['Name'] === $name) {
                return $snapshot;
            }
        }

        throw new \RuntimeException('Snapshot was not found: ' . $name);
    }

    private static function ReadWorkbookSnapshot($workbookPath)
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
        $relationshipTypes = array();
        if ($relationshipsXml !== false) {
            $relationshipsDocument = XDocument::Load($relationshipsXml);
            if ($relationshipsDocument->Root !== null) {
                foreach ($relationshipsDocument->Root->Elements('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship') as $relationship) {
                    $id = (string)$relationship->Attribute('Id');
                    $target = (string)$relationship->Attribute('Target');
                    $type = (string)$relationship->Attribute('Type');
                    if ($id !== '' && $target !== '') {
                        $relationships[$id] = self::ResolvePartUri('/xl/workbook.xml', $target);
                        $relationshipTypes[$id] = $type;
                    }
                }
            }
        }

        $sharedStrings = self::LoadSharedStrings($workbookPath, $relationships, $relationshipTypes);

        $snapshots = array();
        if ($workbookDocument->Root !== null) {
            foreach ($workbookDocument->Root->Elements(self::MainNs . 'sheets') as $sheetsElement) {
                foreach ($sheetsElement->Elements(self::MainNs . 'sheet') as $sheet) {
                    $relationshipId = (string)$sheet->Attribute(self::RelationshipNs . 'id');
                    if ($relationshipId === '' || !array_key_exists($relationshipId, $relationships)) {
                        continue;
                    }

                    $worksheetUri = ltrim($relationships[$relationshipId], '/');

                    $archive = new \ZipArchive();
                    if ($archive->open($workbookPath) !== true) {
                        throw new \RuntimeException('The package could not be opened.');
                    }

                    $worksheetXml = $archive->getFromName($worksheetUri);
                    $archive->close();
                    if ($worksheetXml === false) {
                        continue;
                    }

                    $worksheetDocument = XDocument::Load($worksheetXml);
                    if ($worksheetDocument->Root === null) {
                        continue;
                    }

                    $snapshots[] = self::ReadWorksheetSnapshot($sheet, $worksheetDocument->Root, $sharedStrings);
                }
            }
        }

        return $snapshots;
    }

    private static function ReadWorksheetSnapshot($sheetElement, $worksheetRoot, $sharedStrings)
    {
        $nameAttribute = $sheetElement->Attribute('name');
        $sheetName = $nameAttribute === null ? '' : (string)$nameAttribute->Value;

        $stateAttribute = $sheetElement->Attribute('state');
        $state = $stateAttribute === null ? '' : (string)$stateAttribute->Value;
        if ($state === 'hidden') {
            $visibility = 'Hidden';
        } elseif ($state === 'veryHidden') {
            $visibility = 'VeryHidden';
        } else {
            $visibility = 'Visible';
        }

        $rows = array();
        $sheetData = $worksheetRoot->Element(self::MainNs . 'sheetData');
        $rowElements = $sheetData === null ? array() : $sheetData->Elements(self::MainNs . 'row');
        foreach ($rowElements as $row) {
            $rowIndexAttribute = $row->Attribute('r');
            $rowIndexText = $rowIndexAttribute === null ? '' : (string)$rowIndexAttribute->Value;
            $rowIndex = self::TryParseInt($rowIndexText);
            if ($rowIndex === null) {
                continue;
            }

            $heightAttribute = $row->Attribute('ht');
            $hiddenAttribute = $row->Attribute('hidden');
            $rows[$rowIndex - 1] = array(
                'Height' => self::TryParseDouble($heightAttribute === null ? null : (string)$heightAttribute->Value),
                'Hidden' => self::ParseBool($hiddenAttribute === null ? null : (string)$hiddenAttribute->Value),
            );
        }

        $columns = array();
        $colsElement = $worksheetRoot->Element(self::MainNs . 'cols');
        $colElements = $colsElement === null ? array() : $colsElement->Elements(self::MainNs . 'col');
        foreach ($colElements as $column) {
            $minAttribute = $column->Attribute('min');
            $maxAttribute = $column->Attribute('max');
            $minColumnIndex = self::TryParseInt($minAttribute === null ? null : (string)$minAttribute->Value);
            $maxColumnIndex = self::TryParseInt($maxAttribute === null ? null : (string)$maxAttribute->Value);
            if ($minColumnIndex === null || $maxColumnIndex === null) {
                continue;
            }

            $widthAttribute = $column->Attribute('width');
            $hiddenAttribute = $column->Attribute('hidden');
            $columns[] = array(
                'MinColumnIndex' => $minColumnIndex - 1,
                'MaxColumnIndex' => $maxColumnIndex - 1,
                'Width' => self::TryParseDouble($widthAttribute === null ? null : (string)$widthAttribute->Value),
                'Hidden' => self::ParseBool($hiddenAttribute === null ? null : (string)$hiddenAttribute->Value),
            );
        }

        $merges = array();
        $mergeCellsElement = $worksheetRoot->Element(self::MainNs . 'mergeCells');
        $mergeCellElements = $mergeCellsElement === null ? array() : $mergeCellsElement->Elements(self::MainNs . 'mergeCell');
        foreach ($mergeCellElements as $mergeCell) {
            $refAttribute = $mergeCell->Attribute('ref');
            $merges[] = strtoupper($refAttribute === null ? '' : (string)$refAttribute->Value);
        }

        $cells = array();
        foreach ($worksheetRoot->DescendantsAndSelf() as $cell) {
            if (!($cell instanceof \Aspose\Cells_FOSS\Internal\Xml\XElement) || $cell->Name !== self::MainNs . 'c') {
                continue;
            }

            $cellReferenceAttribute = $cell->Attribute('r');
            $cellReference = strtoupper($cellReferenceAttribute === null ? '' : (string)$cellReferenceAttribute->Value);
            if ($cellReference === '') {
                continue;
            }

            $cells[$cellReference] = self::ReadCellValue($cell, $sharedStrings);
        }

        $sheetViewsElement = $worksheetRoot->Element(self::MainNs . 'sheetViews');
        $sheetView = $sheetViewsElement === null ? null : $sheetViewsElement->Element(self::MainNs . 'sheetView');
        $sheetProtection = $worksheetRoot->Element(self::MainNs . 'sheetProtection');

        $dimensionElement = $worksheetRoot->Element(self::MainNs . 'dimension');
        $dimensionAttribute = $dimensionElement === null ? null : $dimensionElement->Attribute('ref');

        $sheetPrElement = $worksheetRoot->Element(self::MainNs . 'sheetPr');
        $tabColorElement = $sheetPrElement === null ? null : $sheetPrElement->Element(self::MainNs . 'tabColor');
        $tabColorAttribute = $tabColorElement === null ? null : $tabColorElement->Attribute('rgb');

        return array(
            'Name' => $sheetName,
            'Visibility' => $visibility,
            'Dimension' => $dimensionAttribute === null ? '' : (string)$dimensionAttribute->Value,
            'TabColor' => strtoupper($tabColorAttribute === null ? '' : (string)$tabColorAttribute->Value),
            'ShowGridLines' => self::ParseViewBool($sheetView === null ? null : self::AttributeValue($sheetView, 'showGridLines'), true),
            'ShowRowColumnHeaders' => self::ParseViewBool($sheetView === null ? null : self::AttributeValue($sheetView, 'showRowColHeaders'), true),
            'ShowZeros' => self::ParseViewBool($sheetView === null ? null : self::AttributeValue($sheetView, 'showZeros'), true),
            'RightToLeft' => self::ParseViewBool($sheetView === null ? null : self::AttributeValue($sheetView, 'rightToLeft'), false),
            'Zoom' => self::TryParseInt($sheetView === null ? null : self::AttributeValue($sheetView, 'zoomScale')) ?? 100,
            'IsProtected' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'sheet'), $sheetProtection !== null),
            'ProtectObjects' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'objects'), false),
            'ProtectScenarios' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'scenarios'), false),
            'ProtectFormatCells' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'formatCells'), false),
            'ProtectInsertRows' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'insertRows'), false),
            'ProtectAutoFilter' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'autoFilter'), false),
            'ProtectSelectLockedCells' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'selectLockedCells'), false),
            'ProtectSelectUnlockedCells' => self::ParseProtectionBool($sheetProtection === null ? null : self::AttributeValue($sheetProtection, 'selectUnlockedCells'), false),
            'Rows' => $rows,
            'Columns' => $columns,
            'MergeReferences' => $merges,
            'Cells' => $cells,
        );
    }

    private static function CreateOpenXmlWorkbook($path, array $sheets)
    {
        $sheetElements = array();
        $worksheetEntries = array();
        $relationships = array();
        for ($index = 0; $index < count($sheets); $index++) {
            $sheet = $sheets[$index];
            $worksheetEntries['xl/worksheets/sheet' . ($index + 1) . '.xml'] = $sheet['WorksheetXml'];
            $relationships[] = '<Relationship Id="rId' . ($index + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . ($index + 1) . '.xml"/>';
            $stateAttribute = ($sheet['State'] === null || trim($sheet['State']) === '') ? '' : ' state="' . $sheet['State'] . '"';
            $sheetElements[] = '<sheet name="' . $sheet['Name'] . '" sheetId="' . ($index + 1) . '" r:id="rId' . ($index + 1) . '"' . $stateAttribute . '/>';
        }

        $workbookXml = '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' . implode('', $sheetElements) . '</sheets></workbook>';

        $contentTypesXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        for ($index = 0; $index < count($sheets); $index++) {
            $contentTypesXml .= '<Override PartName="/xl/worksheets/sheet' . ($index + 1) . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $contentTypesXml .= '</Types>';

        $rootRelsXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

        $workbookRelsXml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' . implode('', $relationships) . '</Relationships>';

        $entries = array(
            '[Content_Types].xml' => $contentTypesXml,
            '_rels/.rels' => $rootRelsXml,
            'xl/workbook.xml' => $workbookXml,
            'xl/_rels/workbook.xml.rels' => $workbookRelsXml,
        );
        foreach ($worksheetEntries as $entryPath => $entryContent) {
            $entries[$entryPath] = $entryContent;
        }

        ZipPackageHelper::CreatePackage($path, $entries);
    }

    private static function MinimalWorksheetXml()
    {
        return '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
    }

    private static function LoadSharedStrings($packagePath, $relationships, $relationshipTypes)
    {
        $sharedStringsUri = null;
        foreach ($relationshipTypes as $id => $type) {
            if (str_ends_with($type, '/sharedStrings')) {
                $sharedStringsUri = ltrim($relationships[$id], '/');
                break;
            }
        }

        if ($sharedStringsUri === null) {
            return array();
        }

        $archive = new \ZipArchive();
        if ($archive->open($packagePath) !== true) {
            throw new \RuntimeException('The package could not be opened.');
        }

        $sharedStringsXml = $archive->getFromName($sharedStringsUri);
        $archive->close();
        if ($sharedStringsXml === false) {
            return array();
        }

        $document = XDocument::Load($sharedStringsXml);
        $values = array();
        if ($document->Root === null) {
            return $values;
        }

        foreach ($document->Root->Elements(self::MainNs . 'si') as $item) {
            $values[] = self::ReadSharedStringItem($item);
        }

        return $values;
    }

    private static function ReadSharedStringItem($item)
    {
        $textElement = $item->Element(self::MainNs . 't');
        if ($textElement !== null) {
            return $textElement->Value ?? '';
        }

        $result = '';
        foreach ($item->DescendantsAndSelf() as $text) {
            if ($text instanceof \Aspose\Cells_FOSS\Internal\Xml\XElement && $text->Name === self::MainNs . 't') {
                $result .= $text->Value;
            }
        }

        return $result;
    }

    private static function ReadCellValue($cell, $sharedStrings)
    {
        $valueElement = $cell->Element(self::MainNs . 'v');
        $rawValue = $valueElement === null ? '' : (string)$valueElement->Value;
        $typeAttribute = $cell->Attribute('t');
        $type = $typeAttribute === null ? '' : trim((string)$typeAttribute->Value);

        if ($type === 's') {
            $index = self::TryParseInt($rawValue);
            if ($index !== null && $index >= 0 && $index < count($sharedStrings)) {
                return $sharedStrings[$index];
            }
        }

        if ($type === 'inlineStr') {
            $isElement = $cell->Element(self::MainNs . 'is');
            if ($isElement === null) {
                return '';
            }

            $result = '';
            foreach ($isElement->DescendantsAndSelf() as $text) {
                if ($text instanceof \Aspose\Cells_FOSS\Internal\Xml\XElement && $text->Name === self::MainNs . 't') {
                    $result .= (string)$text->Value;
                }
            }

            return $result;
        }

        if ($type === 'b') {
            return $rawValue === '1' ? 'TRUE' : 'FALSE';
        }

        return $rawValue;
    }

    private static function AttributeValue($element, $name)
    {
        $attribute = $element->Attribute($name);
        return $attribute === null ? null : (string)$attribute->Value;
    }

    private static function TryParseDouble($value)
    {
        if ($value === null || !is_numeric($value)) {
            return null;
        }

        return floatval($value);
    }

    private static function TryParseInt($value)
    {
        if ($value === null || !is_numeric($value)) {
            return null;
        }

        return intval($value);
    }

    private static function ParseBool($value)
    {
        if ($value === null) {
            return false;
        }

        return $value === '1' || strcasecmp($value, 'true') === 0;
    }

    private static function ParseViewBool($value, $defaultValue)
    {
        if ($value === null || trim($value) === '') {
            return $defaultValue;
        }

        return $value === '1' || strcasecmp($value, 'true') === 0;
    }

    private static function ParseProtectionBool($value, $defaultValue)
    {
        if ($value === null || trim($value) === '') {
            return $defaultValue;
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
