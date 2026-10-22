<?php

namespace Aspose\Cells_FOSS\MalformedTests;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';
require_once __DIR__ . '/../Shared/AutoFilterScenarioFactory.php';
require_once __DIR__ . '/../Shared/ChartScenarioFactory.php';
require_once __DIR__ . '/../Shared/ShapeScenarioFactory.php';
require_once __DIR__ . '/../Shared/WorkbookMetadataScenarioFactory.php';

use Aspose\Cells_FOSS\BorderStyleType;
use Aspose\Cells_FOSS\CellValueType;
use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\InvalidFileFormatException;
use Aspose\Cells_FOSS\LoadOptions;
use Aspose\Cells_FOSS\PageOrientationType;
use Aspose\Cells_FOSS\Testing\AssertEx;
use Aspose\Cells_FOSS\Testing\AutoFilterScenarioFactory;
use Aspose\Cells_FOSS\Testing\ConditionalFormattingScenarioFactory;
use Aspose\Cells_FOSS\Testing\DefinedNameScenarioFactory;
use Aspose\Cells_FOSS\Testing\HyperlinkScenarioFactory;
use Aspose\Cells_FOSS\Testing\PageSetupScenarioFactory;
use Aspose\Cells_FOSS\Testing\StyleScenarioFactory;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\ValidationScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorkbookMetadataScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorkbookScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorksheetScenarioFactory;
use Aspose\Cells_FOSS\Testing\ZipPackageHelper;
use Aspose\Cells_FOSS\ValidationType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    const MainNs = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}";

    public static function Main()
    {
        return TestRunner::Run(
            'CellData.MalformedTests',
            array(
                new TestCase('missing_workbook_xml_throws_invalid_file_format', array(Program::class, 'MissingWorkbookXmlThrowsInvalidFileFormat')),
                new TestCase('invalid_shared_string_index_records_lossy_recovery', array(Program::class, 'InvalidSharedStringIndexRecordsLossyRecovery')),
                new TestCase('invalid_cell_reference_is_skipped_with_warning', array(Program::class, 'InvalidCellReferenceIsSkippedWithWarning')),
                new TestCase('invalid_cell_reference_throws_in_strict_mode', array(Program::class, 'InvalidCellReferenceThrowsInStrictMode')),
                new TestCase('missing_sheetdata_synthesizes_empty_sheet', array(Program::class, 'MissingSheetDataSynthesizesEmptySheet')),
                new TestCase('missing_workbook_relationships_fall_back_by_convention', array(Program::class, 'MissingWorkbookRelationshipsFallBackByConvention')),
                new TestCase('invalid_date_serial_is_repaired_and_cleared', array(Program::class, 'InvalidDateSerialIsRepairedAndCleared')),
                new TestCase('invalid_style_index_records_warning_and_defaults_to_style_zero', array(Program::class, 'InvalidStyleIndexRecordsWarningAndDefaultsToStyleZero')),
                new TestCase('invalid_style_attributes_fall_back_to_safe_defaults', array(Program::class, 'InvalidStyleAttributesFallBackToSafeDefaults')),
                new TestCase('missing_dimension_records_recovery_and_preserves_settings', array(Program::class, 'MissingDimensionRecordsRecoveryAndPreservesSettings')),
                new TestCase('overlapping_merges_drop_conflicts_with_warning', array(Program::class, 'OverlappingMergesDropConflictsWithWarning')),
                new TestCase('invalid_column_span_throws_invalid_file_format', array(Program::class, 'InvalidColumnSpanThrowsInvalidFileFormat')),
                new TestCase('invalid_hyperlink_ref_is_dropped_with_warning', array(Program::class, 'InvalidHyperlinkRefIsDroppedWithWarning')),
                new TestCase('missing_hyperlink_relationship_records_warning', array(Program::class, 'MissingHyperlinkRelationshipRecordsWarning')),
                new TestCase('invalid_validation_sqref_is_dropped_with_warning', array(Program::class, 'InvalidValidationSqrefIsDroppedWithWarning')),
                new TestCase('invalid_validation_sqref_throws_in_strict_mode', array(Program::class, 'InvalidValidationSqrefThrowsInStrictMode')),
                new TestCase('overlapping_validations_keep_first', array(Program::class, 'OverlappingValidationsKeepFirst')),
                new TestCase('invalid_validation_type_falls_back_with_warning', array(Program::class, 'InvalidValidationTypeFallsBackWithWarning')),
                new TestCase('invalid_conditional_formatting_sqref_is_dropped_with_warning', array(Program::class, 'InvalidConditionalFormattingSqrefIsDroppedWithWarning')),
                new TestCase('unsupported_conditional_formatting_rule_is_dropped_with_warning', array(Program::class, 'UnsupportedConditionalFormattingRuleIsDroppedWithWarning')),
                new TestCase('invalid_conditional_formatting_dxfid_defaults_style', array(Program::class, 'InvalidConditionalFormattingDxfIdDefaultsStyle')),
                new TestCase('invalid_page_setup_attributes_record_warnings_and_fall_back', array(Program::class, 'InvalidPageSetupAttributesRecordWarningsAndFallBack')),
                new TestCase('invalid_sheet_view_attributes_record_warnings_and_fall_back', array(Program::class, 'InvalidSheetViewAttributesRecordWarningsAndFallBack')),
                new TestCase('invalid_sheet_protection_attributes_record_warnings_and_fall_back', array(Program::class, 'InvalidSheetProtectionAttributesRecordWarningsAndFallBack')),
                new TestCase('invalid_autofilter_ref_is_dropped_with_warning', array(Program::class, 'InvalidAutoFilterRefIsDroppedWithWarning')),
                new TestCase('invalid_autofilter_sortstate_records_warning', array(Program::class, 'InvalidAutoFilterSortStateRecordsWarning')),
                new TestCase('invalid_autofilter_dxfids_record_warning_and_drop_invalid_metadata', array(Program::class, 'InvalidAutoFilterDxfIdsRecordWarningAndDropInvalidMetadata')),
                new TestCase('invalid_defined_name_scope_records_warning', array(Program::class, 'InvalidDefinedNameScopeRecordsWarning')),
                new TestCase('invalid_print_titles_defined_name_records_warning', array(Program::class, 'InvalidPrintTitlesDefinedNameRecordsWarning')),
                new TestCase('invalid_workbook_metadata_attributes_record_warning', array(Program::class, 'InvalidWorkbookMetadataAttributesRecordWarning')),
                new TestCase('invalid_document_properties_record_warning', array(Program::class, 'InvalidDocumentPropertiesRecordWarning')),
            )
        );
    }

    public static function MissingWorkbookXmlThrowsInvalidFileFormat()
    {
        $temp = new TemporaryDirectory("malformed-missing-workbook");
        $path = $temp->GetPath("broken.xlsx");
        ZipPackageHelper::CreatePackage($path, array());

        AssertEx::Throws(InvalidFileFormatException::class, function () use ($path) {
            $unused = new Workbook($path);
        });
    }

    public static function InvalidSharedStringIndexRecordsLossyRecovery()
    {
        $temp = new TemporaryDirectory("malformed-sst");
        $path = $temp->GetPath("sst.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue("Hello");
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("<v>0</v>", "<v>99</v>", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal("", $loaded->Worksheets[0]->Cells["A1"]->StringValue);
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "SST-L001"));
    }

    public static function InvalidCellReferenceIsSkippedWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-ref-warning");
        $path = $temp->GetPath("invalid-ref.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue(42);
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("r=\"A1\"", "r=\"1A\"", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(CellValueType::IsNull, $loaded->Worksheets[0]->Cells["A1"]->Type);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "CELL-F001"));
    }

    public static function InvalidCellReferenceThrowsInStrictMode()
    {
        $temp = new TemporaryDirectory("malformed-ref-strict");
        $path = $temp->GetPath("invalid-ref-strict.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue(42);
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("r=\"A1\"", "r=\"1A\"", $xml);
        });

        AssertEx::Throws(InvalidFileFormatException::class, function () use ($path) {
            $options = new LoadOptions();
            $options->StrictMode = true;
            $unused = new Workbook($path, $options);
        });
    }

    public static function MissingSheetDataSynthesizesEmptySheet()
    {
        $temp = new TemporaryDirectory("malformed-sheetdata");
        $path = $temp->GetPath("missing-sheetdata.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue(42);
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $sheetData = $document->Root->Element($mainNs . "sheetData");
            if ($sheetData !== null) {
                $sheetData->Remove();
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(CellValueType::IsNull, $loaded->Worksheets[0]->Cells["A1"]->Type);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "ACF-WS-001"));
    }

    public static function MissingWorkbookRelationshipsFallBackByConvention()
    {
        $temp = new TemporaryDirectory("malformed-rels");
        $path = $temp->GetPath("missing-rels.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue(42);
        $workbook->Save($path);

        ZipPackageHelper::DeleteEntry($path, "xl/_rels/workbook.xml.rels");

        $loaded = new Workbook($path);
        AssertEx::Equal("42", $loaded->Worksheets[0]->Cells["A1"]->StringValue);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "PKG-R001"));
    }

    public static function InvalidDateSerialIsRepairedAndCleared()
    {
        $temp = new TemporaryDirectory("malformed-date");
        $path = $temp->GetPath("invalid-date.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue(new \DateTime("2024-05-06 07:08:09"));
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            $start = strpos($xml, "<v>");
            $end = strpos($xml, "</v>");
            return $start !== false && $end !== false && $end > $start
                ? substr($xml, 0, $start + 3) . "oops" . substr($xml, $end)
                : $xml;
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(CellValueType::IsNull, $loaded->Worksheets[0]->Cells["A1"]->Type);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "CELL-R002"));
    }

    public static function InvalidStyleIndexRecordsWarningAndDefaultsToStyleZero()
    {
        $temp = new TemporaryDirectory("malformed-style-index");
        $path = $temp->GetPath("invalid-style.xlsx");
        $workbook = StyleScenarioFactory::CreateStyledWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("s=\"1\"", "s=\"999\"", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal("1234.567", $loaded->Worksheets[0]->Cells["A1"]->StringValue);
        $style = $loaded->Worksheets[0]->Cells["A1"]->GetStyle();
        AssertEx::False($style->Font->IsBold);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "STYLE-F001"));
    }

    public static function InvalidStyleAttributesFallBackToSafeDefaults()
    {
        $temp = new TemporaryDirectory("malformed-style-attributes");
        $path = $temp->GetPath("invalid-style-attributes.xlsx");
        $workbook = StyleScenarioFactory::CreateStyledWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/styles.xml", function ($document) use ($mainNs) {
            $cellXfs = $document->Root->Element($mainNs . "cellXfs");
            $xfList = $cellXfs === null ? array() : $cellXfs->Elements($mainNs . "xf");
            $xf = count($xfList) > 1 ? $xfList[1] : null;
            if ($xf !== null) {
                $alignment = $xf->Element($mainNs . "alignment");
                if ($alignment !== null) {
                    $alignment->SetAttributeValue("indent", "-5");
                    $alignment->SetAttributeValue("textRotation", "999");
                    $alignment->SetAttributeValue("readingOrder", "42");
                }
            }

            $borders = $document->Root->Element($mainNs . "borders");
            $borderList = $borders === null ? array() : $borders->Elements($mainNs . "border");
            $border = count($borderList) > 1 ? $borderList[1] : null;
            if ($border !== null) {
                $left = $border->Element($mainNs . "left");
                if ($left !== null) {
                    $left->SetAttributeValue("style", "mystery");
                }
            }
        });

        $loaded = new Workbook($path);
        $style = $loaded->Worksheets[0]->Cells["A1"]->GetStyle();
        AssertEx::Equal(0, $style->IndentLevel);
        AssertEx::Equal(0, $style->TextRotation);
        AssertEx::Equal(0, $style->ReadingOrder);
        AssertEx::Equal(BorderStyleType::None, $style->Borders->Left->LineStyle);
    }

    public static function MissingDimensionRecordsRecoveryAndPreservesSettings()
    {
        $temp = new TemporaryDirectory("malformed-missing-dimension");
        $path = $temp->GetPath("missing-dimension.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $dimension = $document->Root->Element($mainNs . "dimension");
            if ($dimension !== null) {
                $dimension->Remove();
            }
        });

        $loaded = new Workbook($path);
        WorksheetScenarioFactory::AssertWorksheetSettings($loaded);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-R001"));
    }

    public static function OverlappingMergesDropConflictsWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-overlapping-merges");
        $path = $temp->GetPath("overlap-merges.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells->Merge(0, 0, 2, 2);
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("</mergeCells>", "<mergeCell ref=\"B2:C3\"/></mergeCells>", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(1, count($loaded->Worksheets[0]->Cells->MergedCells));
        AssertEx::Equal(0, $loaded->Worksheets[0]->Cells->MergedCells[0]->StartRow);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "MRG-L001"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidColumnSpanThrowsInvalidFileFormat()
    {
        $temp = new TemporaryDirectory("malformed-column-span");
        $path = $temp->GetPath("invalid-column-span.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("min=\"1\" max=\"1\"", "min=\"5\" max=\"2\"", $xml);
        });

        AssertEx::Throws(InvalidFileFormatException::class, function () use ($path) {
            $unused = new Workbook($path);
        });
    }

    public static function InvalidHyperlinkRefIsDroppedWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-hyperlink-ref");
        $path = $temp->GetPath("invalid-hyperlink-ref.xlsx");
        $workbook = HyperlinkScenarioFactory::CreateHyperlinkWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $hyperlinks = $document->Root->Element($mainNs . "hyperlinks");
            $hyperlinkList = $hyperlinks === null ? array() : $hyperlinks->Elements($mainNs . "hyperlink");
            if (count($hyperlinkList) > 0) {
                $hyperlinkList[0]->SetAttributeValue("ref", "1A");
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(2, count($loaded->Worksheets[0]->Hyperlinks));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "HL-L001"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function MissingHyperlinkRelationshipRecordsWarning()
    {
        $temp = new TemporaryDirectory("malformed-hyperlink-rels");
        $path = $temp->GetPath("missing-hyperlink-rels.xlsx");
        $workbook = new Workbook();
        $workbook->Worksheets[0]->Cells["A1"]->PutValue("Docs");
        $workbook->Worksheets[0]->Hyperlinks->Add("A1", 1, 1, "https://example.com/docs");
        $workbook->Save($path);

        ZipPackageHelper::DeleteEntry($path, "xl/worksheets/_rels/sheet1.xml.rels");

        $loaded = new Workbook($path);
        AssertEx::Equal(0, count($loaded->Worksheets[0]->Hyperlinks));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "HL-L002"));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "HL-L004"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidValidationSqrefIsDroppedWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-validation-sqref");
        $path = $temp->GetPath("invalid-validation-sqref.xlsx");
        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $dataValidations = $document->Root->Element($mainNs . "dataValidations");
            $validationList = $dataValidations === null ? array() : $dataValidations->Elements($mainNs . "dataValidation");
            if (count($validationList) > 0) {
                $validationList[0]->SetAttributeValue("sqref", "1A");
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(2, count($loaded->Worksheets[0]->Validations));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "DV-L001"));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "DV-L002"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidValidationSqrefThrowsInStrictMode()
    {
        $temp = new TemporaryDirectory("malformed-validation-sqref-strict");
        $path = $temp->GetPath("invalid-validation-sqref-strict.xlsx");
        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("sqref=\"A1:A3\"", "sqref=\"1A\"", $xml);
        });

        AssertEx::Throws(InvalidFileFormatException::class, function () use ($path) {
            $options = new LoadOptions();
            $options->StrictMode = true;
            $unused = new Workbook($path, $options);
        });
    }

    public static function OverlappingValidationsKeepFirst()
    {
        $temp = new TemporaryDirectory("malformed-validation-overlap");
        $path = $temp->GetPath("overlap-validations.xlsx");
        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            $marker = "</dataValidations>";
            $injected = "<dataValidation type=\"whole\" sqref=\"A2:A4\"><formula1>1</formula1><formula2>2</formula2></dataValidation>";
            return str_replace($marker, $injected . $marker, $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(3, count($loaded->Worksheets[0]->Validations));
        AssertEx::Equal(ValidationType::List, $loaded->Worksheets[0]->Validations->GetValidationInCell(1, 0)->Type);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "DV-L003"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidValidationTypeFallsBackWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-validation-type");
        $path = $temp->GetPath("invalid-validation-type.xlsx");
        $workbook = ValidationScenarioFactory::CreateValidationWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("type=\"list\"", "type=\"mystery\"", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(ValidationType::AnyValue, $loaded->Worksheets[0]->Validations[0]->Type);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "DV-R001"));
    }

    public static function InvalidConditionalFormattingSqrefIsDroppedWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-conditional-formatting-sqref");
        $path = $temp->GetPath("invalid-conditional-formatting-sqref.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("sqref=\"A1:A5\"", "sqref=\"1A\"", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(1, count($loaded->Worksheets[0]->ConditionalFormattings));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "CF-L001"));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "CF-L003"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function UnsupportedConditionalFormattingRuleIsDroppedWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-conditional-formatting-type");
        $path = $temp->GetPath("unsupported-conditional-formatting-type.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("type=\"expression\"", "type=\"containsBlanks\"", $xml);
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(2, count($loaded->Worksheets[0]->ConditionalFormattings));
        AssertEx::Equal(1, count($loaded->Worksheets[0]->ConditionalFormattings[0]));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "CF-L002"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidConditionalFormattingDxfIdDefaultsStyle()
    {
        $temp = new TemporaryDirectory("malformed-conditional-formatting-dxfid");
        $path = $temp->GetPath("invalid-conditional-formatting-dxfid.xlsx");
        $workbook = ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteEntryText($path, "xl/worksheets/sheet1.xml", function ($xml) {
            return str_replace("dxfId=\"0\"", "dxfId=\"999\"", $xml);
        });

        $loaded = new Workbook($path);
        $collection = $loaded->Worksheets[0]->ConditionalFormattings[0];
        AssertEx::Equal(FillPattern::None, $collection[0]->Style->Pattern);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "CF-R002"));
    }

    public static function InvalidPageSetupAttributesRecordWarningsAndFallBack()
    {
        $temp = new TemporaryDirectory("malformed-page-setup");
        $path = $temp->GetPath("page-setup-invalid.xlsx");
        $workbook = PageSetupScenarioFactory::CreatePageSetupWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $xmlPageSetup = $document->Root->Element($mainNs . "pageSetup");
            if ($xmlPageSetup !== null) {
                $xmlPageSetup->SetAttributeValue("scale", "999");
                $xmlPageSetup->SetAttributeValue("orientation", "sideways");
            }
            $pageMargins = $document->Root->Element($mainNs . "pageMargins");
            if ($pageMargins !== null) {
                $pageMargins->SetAttributeValue("left", "-1");
            }
        });

        $loaded = new Workbook($path);
        $pageSetup = $loaded->Worksheets[0]->PageSetup;
        AssertEx::Equal(95, $pageSetup->Scale === null ? 95 : $pageSetup->Scale);
        AssertEx::Equal(PageOrientationType::Default, $pageSetup->Orientation);
        AssertEx::Equal(0.7, $pageSetup->LeftMarginInch);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "PG-L001"));
        AssertEx::True(self::CountIssues($loaded->LoadDiagnostics, "PG-L002") >= 1);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "PG-L003"));
    }

    public static function InvalidSheetProtectionAttributesRecordWarningsAndFallBack()
    {
        $temp = new TemporaryDirectory("malformed-sheet-protection");
        $path = $temp->GetPath("invalid-sheet-protection.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $protection = $document->Root->Element($mainNs . "sheetProtection");
            if ($protection !== null) {
                $protection->SetAttributeValue("sheet", "maybe");
                $protection->SetAttributeValue("objects", "oops");
                $protection->SetAttributeValue("formatCells", "1");
                $protection->SetAttributeValue("selectUnlockedCells", "1");
            }
        });

        $loaded = new Workbook($path);
        $sheet = $loaded->Worksheets[0];
        AssertEx::True($sheet->Protection->IsProtected);
        AssertEx::False($sheet->Protection->Objects);
        AssertEx::True($sheet->Protection->FormatCells);
        AssertEx::True($sheet->Protection->SelectUnlockedCells);
        AssertEx::True(self::CountIssues($loaded->LoadDiagnostics, "WS-L009") >= 2);
    }

    public static function InvalidAutoFilterRefIsDroppedWithWarning()
    {
        $temp = new TemporaryDirectory("malformed-autofilter-ref");
        $path = $temp->GetPath("invalid-autofilter-ref.xlsx");
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $autoFilter = $document->Root->Element($mainNs . "autoFilter");
            if ($autoFilter !== null) {
                $autoFilter->SetAttributeValue("ref", "1A");
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal("", $loaded->Worksheets[0]->AutoFilter->Range);
        AssertEx::Equal(0, count($loaded->Worksheets[0]->AutoFilter->FilterColumns));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L010"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidAutoFilterSortStateRecordsWarning()
    {
        $temp = new TemporaryDirectory("malformed-autofilter-sortstate");
        $path = $temp->GetPath("invalid-autofilter-sortstate.xlsx");
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $autoFilter = $document->Root->Element($mainNs . "autoFilter");
            if ($autoFilter !== null) {
                $sortState = $autoFilter->Element($mainNs . "sortState");
                if ($sortState !== null) {
                    $sortState->SetAttributeValue("ref", "1A");
                }
                $autoFilter->Add(new XElement($mainNs . "filterColumn", new XAttribute("colId", "bad")));
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal("A1:E6", $loaded->Worksheets[0]->AutoFilter->Range);
        AssertEx::Equal(5, count($loaded->Worksheets[0]->AutoFilter->FilterColumns));
        AssertEx::Equal("", $loaded->Worksheets[0]->AutoFilter->SortState->Ref);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L011"));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L012"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidAutoFilterDxfIdsRecordWarningAndDropInvalidMetadata()
    {
        $temp = new TemporaryDirectory("malformed-autofilter-dxfids");
        $path = $temp->GetPath("invalid-autofilter-dxfids.xlsx");
        $workbook = AutoFilterScenarioFactory::CreateAutoFilterWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $autoFilter = $document->Root->Element($mainNs . "autoFilter");
            if ($autoFilter !== null) {
                foreach ($autoFilter->Elements($mainNs . "filterColumn") as $filterColumn) {
                    $colIdAttribute = $filterColumn->Attribute("colId");
                    $colId = $colIdAttribute === null ? null : (string) $colIdAttribute;
                    if ($colId === "2") {
                        $colorFilter = $filterColumn->Element($mainNs . "colorFilter");
                        if ($colorFilter !== null) {
                            $colorFilter->SetAttributeValue("dxfId", "99");
                        }
                    }
                }

                $sortState = $autoFilter->Element($mainNs . "sortState");
                if ($sortState !== null) {
                    $sortCondition = $sortState->Element($mainNs . "sortCondition");
                    if ($sortCondition !== null) {
                        $sortCondition->SetAttributeValue("dxfId", "99");
                    }
                }
            }
        });

        $loaded = new Workbook($path);
        $colorColumn = $loaded->Worksheets[0]->AutoFilter->FilterColumns[2];
        AssertEx::False($colorColumn->ColorFilter->Enabled);
        AssertEx::Null($colorColumn->ColorFilter->DifferentialStyleId);
        AssertEx::True(count($loaded->Worksheets[0]->AutoFilter->SortState->SortConditions) > 0);
        AssertEx::Null($loaded->Worksheets[0]->AutoFilter->SortState->SortConditions[0]->DifferentialStyleId);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L011"));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L012"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidDefinedNameScopeRecordsWarning()
    {
        $temp = new TemporaryDirectory("malformed-defined-name-scope");
        $path = $temp->GetPath("invalid-defined-name-scope.xlsx");
        $workbook = DefinedNameScenarioFactory::CreateDefinedNamesWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/workbook.xml", function ($document) use ($mainNs) {
            $definedNames = $document->Root->Element($mainNs . "definedNames");
            if ($definedNames !== null) {
                foreach ($definedNames->Elements($mainNs . "definedName") as $definedName) {
                    $nameAttribute = $definedName->Attribute("name");
                    $name = $nameAttribute === null ? null : (string) $nameAttribute;
                    if ($name === "LocalCell") {
                        $definedName->SetAttributeValue("localSheetId", "99");
                    }
                }
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal(1, count($loaded->DefinedNames));
        AssertEx::Equal("GlobalRange", $loaded->DefinedNames[0]->Name);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WB-L002"));
        AssertEx::True($loaded->LoadDiagnostics->HasDataLossRisk);
    }

    public static function InvalidPrintTitlesDefinedNameRecordsWarning()
    {
        $temp = new TemporaryDirectory("malformed-print-titles");
        $path = $temp->GetPath("invalid-print-titles.xlsx");
        $workbook = PageSetupScenarioFactory::CreatePageSetupWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/workbook.xml", function ($document) use ($mainNs) {
            $definedNames = $document->Root->Element($mainNs . "definedNames");
            if ($definedNames !== null) {
                foreach ($definedNames->Elements($mainNs . "definedName") as $definedName) {
                    $nameAttribute = $definedName->Attribute("name");
                    $name = $nameAttribute === null ? null : (string) $nameAttribute;
                    if ($name === "_xlnm.Print_Titles") {
                        $definedName->Value = '\'Print Sheet\'!$A$1';
                    }
                }
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal("", $loaded->Worksheets[0]->PageSetup->PrintTitleRows);
        AssertEx::Equal("", $loaded->Worksheets[0]->PageSetup->PrintTitleColumns);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "PG-L004"));
    }

    public static function InvalidSheetViewAttributesRecordWarningsAndFallBack()
    {
        $temp = new TemporaryDirectory("malformed-sheet-view");
        $path = $temp->GetPath("invalid-sheet-view.xlsx");
        $workbook = WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/worksheets/sheet1.xml", function ($document) use ($mainNs) {
            $sheetViews = $document->Root->Element($mainNs . "sheetViews");
            if ($sheetViews !== null) {
                $sheetView = $sheetViews->Element($mainNs . "sheetView");
                if ($sheetView !== null) {
                    $sheetView->SetAttributeValue("zoomScale", "999");
                    $sheetView->SetAttributeValue("showGridLines", "maybe");
                }
            }
            $sheetPr = $document->Root->Element($mainNs . "sheetPr");
            if ($sheetPr !== null) {
                $tabColor = $sheetPr->Element($mainNs . "tabColor");
                if ($tabColor !== null) {
                    $tabColor->SetAttributeValue("rgb", "broken");
                }
            }
        });

        $loaded = new Workbook($path);
        $sheet = $loaded->Worksheets[0];
        AssertEx::Equal(Color::Empty(), $sheet->TabColor);
        AssertEx::True($sheet->ShowGridlines);
        AssertEx::False($sheet->ShowRowColumnHeaders);
        AssertEx::False($sheet->ShowZeros);
        AssertEx::True($sheet->RightToLeft);
        AssertEx::Equal(100, $sheet->Zoom);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L007"));
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WS-L008"));
    }

    public static function InvalidWorkbookMetadataAttributesRecordWarning()
    {
        $temp = new TemporaryDirectory("malformed-workbook-metadata");
        $path = $temp->GetPath("workbook-metadata.xlsx");
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($path);

        $mainNs = self::MainNs;
        ZipPackageHelper::RewriteXmlEntry($path, "xl/workbook.xml", function ($document) use ($mainNs) {
            $workbookPr = $document->Root->Element($mainNs . "workbookPr");
            if ($workbookPr !== null) {
                $workbookPr->SetAttributeValue("showObjects", "mystery");
            }

            $bookViews = $document->Root->Element($mainNs . "bookViews");
            if ($bookViews !== null) {
                $workbookView = $bookViews->Element($mainNs . "workbookView");
                if ($workbookView !== null) {
                    $workbookView->SetAttributeValue("activeTab", "99");
                }
            }

            $calcPr = $document->Root->Element($mainNs . "calcPr");
            if ($calcPr !== null) {
                $calcPr->SetAttributeValue("iterateDelta", "oops");
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Equal("all", $loaded->Properties->ShowObjects);
        AssertEx::Equal(0.001, $loaded->Properties->Calculation->IterateDelta);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WB-L003"));
    }

    public static function InvalidDocumentPropertiesRecordWarning()
    {
        $temp = new TemporaryDirectory("malformed-document-properties");
        $path = $temp->GetPath("document-properties.xlsx");
        $workbook = WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook();
        $workbook->Save($path);

        ZipPackageHelper::RewriteXmlEntry($path, "docProps/core.xml", function ($document) {
            $created = $document->Root->Element("{http://purl.org/dc/terms/}created");
            if ($created !== null) {
                $created->Value = "not-a-date";
            }
        });

        $loaded = new Workbook($path);
        AssertEx::Null($loaded->DocumentProperties->Core->Created);
        AssertEx::True(self::HasIssue($loaded->LoadDiagnostics, "WB-L004"));
    }

    private static function HasIssue($diagnostics, $code)
    {
        foreach ($diagnostics->Issues as $issue) {
            if ($issue->Code === $code) {
                return true;
            }
        }

        return false;
    }

    private static function CountIssues($diagnostics, $code)
    {
        $count = 0;
        foreach ($diagnostics->Issues as $issue) {
            if ($issue->Code === $code) {
                $count++;
            }
        }

        return $count;
    }
}

Program::Main();
