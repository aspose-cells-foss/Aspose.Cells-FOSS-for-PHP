<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

require_once __DIR__ . '/../../autoload.php';
require_once __DIR__ . '/../Shared/TestInfrastructure.php';
require_once __DIR__ . '/../Shared/AutoFilterScenarioFactory.php';
require_once __DIR__ . '/../Shared/WorkbookMetadataScenarioFactory.php';
require_once __DIR__ . '/ComparisonValueHelpers.php';
require_once __DIR__ . '/CellSnapshot.php';
require_once __DIR__ . '/StyleSnapshot.php';
require_once __DIR__ . '/SnapshotMismatch.php';
require_once __DIR__ . '/OpenXmlDateSystem.php';
require_once __DIR__ . '/OpenXmlReadContext.php';
require_once __DIR__ . '/OpenXmlCellDescriptor.php';
require_once __DIR__ . '/OpenXmlStylesContext.php';
require_once __DIR__ . '/OpenXmlFontDescriptor.php';
require_once __DIR__ . '/OpenXmlFillDescriptor.php';
require_once __DIR__ . '/OpenXmlBorderSideDescriptor.php';
require_once __DIR__ . '/OpenXmlBorderDescriptor.php';
require_once __DIR__ . '/OpenXmlComparisonSupport.php';

use Aspose\Cells_FOSS\Cell;
use Aspose\Cells_FOSS\CellValueType;
use Aspose\Cells_FOSS\Workbook;
use Aspose\Cells_FOSS\Testing\AutoFilterScenarioFactory;
use Aspose\Cells_FOSS\Testing\ConditionalFormattingScenarioFactory;
use Aspose\Cells_FOSS\Testing\DefinedNameScenarioFactory;
use Aspose\Cells_FOSS\Testing\HyperlinkScenarioFactory;
use Aspose\Cells_FOSS\Testing\PageSetupScenarioFactory;
use Aspose\Cells_FOSS\Testing\StyleScenarioFactory;
use Aspose\Cells_FOSS\Testing\TestCase;
use Aspose\Cells_FOSS\Testing\TestRunner;
use Aspose\Cells_FOSS\Testing\TemporaryDirectory;
use Aspose\Cells_FOSS\Testing\ValidationScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorkbookScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorkbookMetadataScenarioFactory;
use Aspose\Cells_FOSS\Testing\WorksheetScenarioFactory;

final class Program
{
    public static function Main()
    {
        $dataTemp = new TemporaryDirectory('compare-openxml-data');
        $styleTemp = new TemporaryDirectory('compare-openxml-styles');

        $dataLogPath = $dataTemp->GetPath('mismatches.log');
        $styleLogPath = $styleTemp->GetPath('mismatches.log');

        $dataCases = array(
            new TestCase('mixed_cell_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('mixed.xlsx');
                WorkbookScenarioFactory::CreateMixedCellWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('worksheet_settings_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('worksheet-settings.xlsx');
                WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('page_setup_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('page-setup.xlsx');
                PageSetupScenarioFactory::CreatePageSetupWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('hyperlinks_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('hyperlinks.xlsx');
                HyperlinkScenarioFactory::CreateHyperlinkWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('styled_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('styled.xlsx');
                StyleScenarioFactory::CreateStyledWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('validation_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('validation.xlsx');
                ValidationScenarioFactory::CreateValidationWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('conditional_formatting_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('conditional-formatting.xlsx');
                ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('defined_names_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('defined-names.xlsx');
                DefinedNameScenarioFactory::CreateDefinedNamesWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('autofilter_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('autofilter.xlsx');
                AutoFilterScenarioFactory::CreateAutoFilterWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
            new TestCase('workbook_metadata_workbook_data_matches_openxml_sdk', function () use ($dataTemp, $dataLogPath) {
                $path = $dataTemp->GetPath('workbook-metadata.xlsx');
                WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook()->Save($path);
                self::AssertNoDataMismatches($path, $dataLogPath);
            }),
        );

        $styleCases = array(
            new TestCase('mixed_cell_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('mixed.xlsx');
                WorkbookScenarioFactory::CreateMixedCellWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('worksheet_settings_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('worksheet-settings.xlsx');
                WorksheetScenarioFactory::CreateWorksheetSettingsWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('page_setup_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('page-setup.xlsx');
                PageSetupScenarioFactory::CreatePageSetupWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('hyperlinks_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('hyperlinks.xlsx');
                HyperlinkScenarioFactory::CreateHyperlinkWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('styled_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('styled.xlsx');
                StyleScenarioFactory::CreateStyledWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('validation_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('validation.xlsx');
                ValidationScenarioFactory::CreateValidationWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('conditional_formatting_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('conditional-formatting.xlsx');
                ConditionalFormattingScenarioFactory::CreateConditionalFormattingWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('defined_names_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('defined-names.xlsx');
                DefinedNameScenarioFactory::CreateDefinedNamesWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('autofilter_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('autofilter.xlsx');
                AutoFilterScenarioFactory::CreateAutoFilterWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
            new TestCase('workbook_metadata_workbook_styles_match_openxml_sdk', function () use ($styleTemp, $styleLogPath) {
                $path = $styleTemp->GetPath('workbook-metadata.xlsx');
                WorkbookMetadataScenarioFactory::CreateWorkbookMetadataWorkbook()->Save($path);
                self::AssertNoStyleMismatches($path, $styleLogPath);
            }),
        );

        foreach (self::DiscoverCompareFiles(self::CompareInputRoot()) as $filePath) {
            $capturedPath = $filePath;
            $caseName = strtolower(pathinfo($filePath, PATHINFO_FILENAME));
            $caseName = str_replace(' ', '_', $caseName);
            $dataCases[] = new TestCase($caseName . '_data_matches_openxml_sdk', function () use ($capturedPath, $dataLogPath) {
                self::AssertNoDataMismatches($capturedPath, $dataLogPath);
            });
            $styleCases[] = new TestCase($caseName . '_styles_match_openxml_sdk', function () use ($capturedPath, $styleLogPath) {
                self::AssertNoStyleMismatches($capturedPath, $styleLogPath);
            });
        }

        $dataResult = TestRunner::Run('CompareOpenXml.DataTests', $dataCases);
        $styleResult = TestRunner::Run('CompareOpenXml.StyleTests', $styleCases);

        if (file_exists($dataLogPath)) {
            echo 'Data mismatch log: ' . $dataLogPath . PHP_EOL;
        }

        if (file_exists($styleLogPath)) {
            echo 'Style mismatch log: ' . $styleLogPath . PHP_EOL;
        }

        return $dataResult | $styleResult;
    }

    private static function DiscoverCompareFiles($root)
    {
        if (!is_dir($root)) {
            return array();
        }

        $supportedExtensions = array('.xlsx', '.xlsm', '.xltx', '.xltm');
        $files = array();
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($root, \FilesystemIterator::SKIP_DOTS));
        foreach ($iterator as $file) {
            if (!$file->isFile()) {
                continue;
            }

            $extension = strtolower($file->getExtension());
            if (!in_array('.' . $extension, $supportedExtensions, true)) {
                continue;
            }

            if (strpos($file->getBasename(), '~$') === 0) {
                continue;
            }

            $files[] = $file->getPathname();
        }

        usort($files, function ($left, $right) {
            return strcasecmp($left, $right);
        });

        return $files;
    }

    private static function AssertNoDataMismatches($workbookPath, $logPath)
    {
        $mismatches = self::CompareData($workbookPath);
        if (count($mismatches) === 0) {
            return;
        }

        self::AppendMismatchLog($logPath, $workbookPath, $mismatches, 'cell mismatch(es)');

        $message = count($mismatches) . " data mismatch(es) in '" . basename($workbookPath) . "':" . PHP_EOL;
        foreach ($mismatches as $mismatch) {
            $message .= '  Cell: ' . $mismatch->CellKey . PHP_EOL;
            $message .= '    Library: ' . $mismatch->LibrarySnapshot . PHP_EOL;
            $message .= '    OpenXml: ' . $mismatch->OpenXmlSnapshot . PHP_EOL;
        }

        throw new \RuntimeException($message);
    }

    private static function AssertNoStyleMismatches($workbookPath, $logPath)
    {
        $mismatches = self::CompareStyles($workbookPath);
        if (count($mismatches) === 0) {
            return;
        }

        self::AppendMismatchLog($logPath, $workbookPath, $mismatches, 'style mismatch(es)');

        $message = count($mismatches) . " style mismatch(es) in '" . basename($workbookPath) . "':" . PHP_EOL;
        foreach ($mismatches as $mismatch) {
            $message .= '  Cell: ' . $mismatch->CellKey . PHP_EOL;
            $message .= '    Library: ' . $mismatch->LibrarySnapshot . PHP_EOL;
            $message .= '    OpenXml: ' . $mismatch->OpenXmlSnapshot . PHP_EOL;
        }

        throw new \RuntimeException($message);
    }

    private static function AppendMismatchLog($logPath, $workbookPath, $mismatches, $mismatchLabel)
    {
        $text = 'Workbook: ' . basename($workbookPath) . PHP_EOL;
        $text .= 'Path: ' . $workbookPath . PHP_EOL;
        $text .= 'Result: mismatched (' . count($mismatches) . ' ' . $mismatchLabel . ')' . PHP_EOL;
        foreach ($mismatches as $mismatch) {
            $text .= 'Cell: ' . $mismatch->CellKey . PHP_EOL;
            $text .= '  Aspose.Cells_FOSS: ' . $mismatch->LibrarySnapshot . PHP_EOL;
            $text .= '  Open XML SDK:     ' . $mismatch->OpenXmlSnapshot . PHP_EOL;
        }

        $text .= PHP_EOL;
        file_put_contents($logPath, $text, FILE_APPEND);
    }

    private static function CompareData($workbookPath)
    {
        $libraryCells = self::ReadCellDataWithLibrary($workbookPath);
        $openXmlCells = OpenXmlComparisonSupport::ReadCellDataWithOpenXmlSdk($workbookPath);
        return self::CompareCellMaps($libraryCells, $openXmlCells);
    }

    private static function CompareStyles($workbookPath)
    {
        $libraryStyles = self::ReadCellStylesWithLibrary($workbookPath);
        $openXmlStyles = OpenXmlComparisonSupport::ReadCellStylesWithOpenXmlSdk($workbookPath);
        return self::CompareStyleMaps($libraryStyles, $openXmlStyles);
    }

    private static function ReadCellDataWithLibrary($workbookPath)
    {
        $workbook = new Workbook($workbookPath);
        return self::EnumerateLibraryCells($workbook, function ($sheetName, $cellReference, $cell) {
            return new CellSnapshot(
                $sheetName,
                $cellReference,
                self::NormalizeCellType($cell->Type),
                ComparisonValueHelpers::NormalizeValue($cell->Value),
                ComparisonValueHelpers::NormalizeFormulaText($cell->Formula)
            );
        });
    }

    private static function ReadCellStylesWithLibrary($workbookPath)
    {
        $workbook = new Workbook($workbookPath);
        return self::EnumerateLibraryCells($workbook, function ($sheetName, $cellReference, $cell) {
            return OpenXmlComparisonSupport::CreateLibraryStyleSnapshot($sheetName, $cellReference, $cell->GetStyle());
        });
    }

    private static function EnumerateLibraryCells(Workbook $workbook, $snapshotFactory)
    {
        $workbookModel = $workbook->Model;
        if ($workbookModel === null) {
            throw new \RuntimeException('Workbook.Model returned null.');
        }

        $worksheets = $workbookModel->Worksheets;
        if (!is_array($worksheets)) {
            throw new \RuntimeException('WorkbookModel.Worksheets is not enumerable.');
        }

        $map = array();
        foreach ($worksheets as $worksheet) {
            if ($worksheet === null) {
                continue;
            }

            $sheetName = $worksheet->Name;
            if ($sheetName === null || $sheetName === '') {
                $sheetName = '<unknown>';
            }

            $sheet = $workbook->Worksheets[$sheetName];
            $cells = $worksheet->Cells;
            if (!is_array($cells)) {
                throw new \RuntimeException('WorksheetModel.Cells is not enumerable.');
            }

            foreach ($cells as $addressText => $cellRecord) {
                $cellReference = '';
                if (!ComparisonValueHelpers::TryNormalizeCellReference($addressText, $cellReference)) {
                    continue;
                }

                $cell = $sheet->Cells[$cellReference];
                $map[ComparisonValueHelpers::BuildCellKey($sheetName, $cellReference)] = $snapshotFactory($sheetName, $cellReference, $cell);
            }
        }

        return $map;
    }

    private static function CompareCellMaps($libraryCells, $openXmlCells)
    {
        $allKeys = self::UnionKeys($libraryCells, $openXmlCells);
        $mismatches = array();
        foreach ($allKeys as $key) {
            $librarySnapshot = self::Lookup($libraryCells, $key);
            $openXmlSnapshot = self::Lookup($openXmlCells, $key);

            if ($librarySnapshot === null || $openXmlSnapshot === null) {
                if (self::IsEffectivelyBlank($librarySnapshot) && self::IsEffectivelyBlank($openXmlSnapshot)) {
                    continue;
                }

                $mismatches[] = new SnapshotMismatch(
                    $key,
                    $librarySnapshot === null ? '<missing>' : $librarySnapshot->ToString(),
                    $openXmlSnapshot === null ? '<missing>' : $openXmlSnapshot->ToString()
                );
                continue;
            }

            if ($librarySnapshot->CellType !== $openXmlSnapshot->CellType
                || $librarySnapshot->Value !== $openXmlSnapshot->Value
                || $librarySnapshot->Formula !== $openXmlSnapshot->Formula) {
                $mismatches[] = new SnapshotMismatch($key, $librarySnapshot->ToString(), $openXmlSnapshot->ToString());
            }
        }

        return $mismatches;
    }

    private static function CompareStyleMaps($libraryStyles, $openXmlStyles)
    {
        $allKeys = self::UnionKeys($libraryStyles, $openXmlStyles);
        $mismatches = array();
        foreach ($allKeys as $key) {
            $librarySnapshot = self::Lookup($libraryStyles, $key);
            $openXmlSnapshot = self::Lookup($openXmlStyles, $key);

            if ($librarySnapshot === null || $openXmlSnapshot === null) {
                if (self::IsEffectivelyDefaultStyle($librarySnapshot) && self::IsEffectivelyDefaultStyle($openXmlSnapshot)) {
                    continue;
                }

                $mismatches[] = new SnapshotMismatch(
                    $key,
                    $librarySnapshot === null ? '<missing>' : $librarySnapshot->ToString(),
                    $openXmlSnapshot === null ? '<missing>' : $openXmlSnapshot->ToString()
                );
                continue;
            }

            if (!$librarySnapshot->Equals($openXmlSnapshot)) {
                $mismatches[] = new SnapshotMismatch($key, $librarySnapshot->ToString(), $openXmlSnapshot->ToString());
            }
        }

        return $mismatches;
    }

    private static function UnionKeys($leftMap, $rightMap)
    {
        $seen = array();
        $keys = array();
        foreach (array_keys($leftMap) as $key) {
            $upper = strtoupper($key);
            if (!array_key_exists($upper, $seen)) {
                $seen[$upper] = true;
                $keys[] = $key;
            }
        }

        foreach (array_keys($rightMap) as $key) {
            $upper = strtoupper($key);
            if (!array_key_exists($upper, $seen)) {
                $seen[$upper] = true;
                $keys[] = $key;
            }
        }

        usort($keys, function ($left, $right) {
            return strcasecmp($left, $right);
        });

        return $keys;
    }

    private static function Lookup($map, $key)
    {
        if (array_key_exists($key, $map)) {
            return $map[$key];
        }

        foreach ($map as $candidateKey => $value) {
            if (strcasecmp($candidateKey, $key) === 0) {
                return $value;
            }
        }

        return null;
    }

    private static function IsEffectivelyBlank($snapshot)
    {
        if ($snapshot === null) {
            return true;
        }

        return $snapshot->CellType === 'Blank'
            && $snapshot->Formula === ''
            && $snapshot->Value === '<null>';
    }

    private static function IsEffectivelyDefaultStyle($snapshot)
    {
        return $snapshot === null || $snapshot->Equals(StyleSnapshot::Default());
    }

    private static function NormalizeCellType($type)
    {
        switch ($type) {
            case CellValueType::IsString:
                return 'String';
            case CellValueType::IsNumeric:
                return 'Number';
            case CellValueType::IsBool:
                return 'Boolean';
            case CellValueType::IsDateTime:
                return 'DateTime';
            case CellValueType::IsUnknown:
                return 'Formula';
            default:
                return 'Blank';
        }
    }

    private static function CompareInputRoot()
    {
        return self::ResolveRepoRoot() . DIRECTORY_SEPARATOR . 'Input';
    }

    private static function ResolveRepoRoot()
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== '') {
            if (is_file($directory . DIRECTORY_SEPARATOR . 'composer.json')) {
                return $directory;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }

            $directory = $parent;
        }

        return getcwd();
    }
}

Program::Main();
