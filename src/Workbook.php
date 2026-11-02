<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Rendering\WorkbookPdfExporter;

class Workbook
{
    private $_model;
    private $_worksheets;
    private $_settings;
    private $_properties;
    private $_documentProperties;
    private $_definedNames;

    public $LoadDiagnostics;

    public function __construct($fileName = null, $options = null)
    {
        $this->_model = new WorkbookModel();
        $this->_worksheets = new WorksheetCollection($this);
        $this->_settings = new WorkbookSettings($this->_model->Settings);
        $this->_properties = new WorkbookProperties($this->_model);
        $this->_documentProperties = new DocumentProperties($this->_model->DocumentProperties);
        $this->_definedNames = new DefinedNameCollection($this);
        $this->LoadDiagnostics = new LoadDiagnostics();

        if ($fileName === null) {
            return;
        }

        if ($options === null) {
            $options = new LoadOptions();
        }

        $this->LoadFromStream($fileName, $options);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Worksheets') {
            return $this->_worksheets;
        }
        if ($propertyName === 'Settings') {
            return $this->_settings;
        }
        if ($propertyName === 'Properties') {
            return $this->_properties;
        }
        if ($propertyName === 'DocumentProperties') {
            return $this->_documentProperties;
        }
        if ($propertyName === 'DefinedNames') {
            return $this->_definedNames;
        }
        if ($propertyName === 'Model') {
            return $this->_model;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Worksheets') {
            throw new \RuntimeException('Worksheets is read-only.');
        }
        if ($propertyName === 'Settings') {
            throw new \RuntimeException('Settings is read-only.');
        }
        if ($propertyName === 'Properties') {
            throw new \RuntimeException('Properties is read-only.');
        }
        if ($propertyName === 'DocumentProperties') {
            throw new \RuntimeException('DocumentProperties is read-only.');
        }
        if ($propertyName === 'DefinedNames') {
            throw new \RuntimeException('DefinedNames is read-only.');
        }
        if ($propertyName === 'Model') {
            throw new \RuntimeException('Model is read-only.');
        }
    }

    public function EnsureUniqueSheetName($sheetName, $currentSheet = null)
    {
        for ($index = 0; $index < count($this->_model->Worksheets); $index++) {
            $existing = $this->_model->Worksheets[$index];
            if ($existing === $currentSheet) {
                continue;
            }

            if (strcasecmp($existing->Name, $sheetName) === 0) {
                throw new CellsException("Worksheet name '" . $sheetName . "' already exists.");
            }
        }
    }

    public function EnsureValidDefinedNameScope($localSheetIndex)
    {
        if ($localSheetIndex === null) {
            return;
        }

        if ($localSheetIndex < 0 || $localSheetIndex >= count($this->_model->Worksheets)) {
            throw new CellsException('Defined name scope must refer to an existing worksheet.');
        }
    }

    public function EnsureUniqueDefinedName($currentDefinedName, $name, $localSheetIndex)
    {
        for ($index = 0; $index < count($this->_model->DefinedNames); $index++) {
            $existing = $this->_model->DefinedNames[$index];
            if ($existing === $currentDefinedName) {
                continue;
            }

            if (strcasecmp($existing->Name, $name) !== 0) {
                continue;
            }

            if (DefinedNameUtility::SameScope($existing->LocalSheetIndex, $localSheetIndex)) {
                throw new CellsException("Defined name '" . $name . "' already exists in the same scope.");
            }
        }
    }

    public function Save($fileName, $options = null)
    {
        if ($options === null) {
            $options = new SaveOptions();
            if ($fileName !== null && str_ends_with(strtolower($fileName), '.pdf')) {
                $options->SaveFormat = SaveFormat::Pdf;
            }
        } elseif (!($options instanceof SaveOptions)) {
            $saveOptions = new SaveOptions();
            $saveOptions->SaveFormat = $options;
            $options = $saveOptions;
        }

        if ($fileName === null) {
            throw new \InvalidArgumentException('fileName');
        }
        if ($options === null) {
            throw new \InvalidArgumentException('options');
        }

        try {
            if ($options->SaveFormat === SaveFormat::Pdf) {
                $pdfOptions = $options instanceof PdfSaveOptions ? $options : new PdfSaveOptions();
                WorkbookPdfExporter::Export($this->_model, $fileName, $pdfOptions);
            } else {
                XlsxWorkbookSerializer::Save($this->_model, $fileName, $options);
            }
        } catch (CellsException $exception) {
            throw $exception;
        } catch (\Throwable $exception) {
            throw new WorkbookSaveException('Failed to save workbook.', $exception);
        }
    }

    public function Dispose()
    {
    }

    private function LoadFromStream($stream, $options)
    {
        try {
            $loadedModel = XlsxWorkbookSerializer::Load($stream, $options, $this->LoadDiagnostics);
            // Keep the Workbook facade and child facade instances stable after load by
            // copying the loaded state into the existing model graph.
            $this->_model->Worksheets = array();
            $this->_model->Worksheets = array_merge($this->_model->Worksheets, $loadedModel->Worksheets);
            $this->_model->Settings->DateSystem = $loadedModel->Settings->DateSystem;
            $this->_model->Settings->DisplayCulture = $loadedModel->Settings->DisplayCulture->Clone();
            $this->_model->Properties->CopyFrom($loadedModel->Properties);
            $this->_model->DocumentProperties->CopyFrom($loadedModel->DocumentProperties);
            $this->_model->DefaultStyle = $loadedModel->DefaultStyle->Clone();
            $this->_model->ActiveSheetIndex = $loadedModel->ActiveSheetIndex;
            $this->_model->DefinedNames = array();
            $this->_model->DefinedNames = array_merge($this->_model->DefinedNames, $loadedModel->DefinedNames);
            $this->_model->RawThemeXml = $loadedModel->RawThemeXml;
            $this->_model->RawTableStylesXml = $loadedModel->RawTableStylesXml;
            $this->_model->TableDifferentialStyles = array();
            $this->_model->TableDifferentialStyles = array_merge($this->_model->TableDifferentialStyles, $loadedModel->TableDifferentialStyles);
            $this->_model->RawTableDifferentialStylesXml = array();
            $this->_model->RawTableDifferentialStylesXml = array_merge($this->_model->RawTableDifferentialStylesXml, $loadedModel->RawTableDifferentialStylesXml);
            $this->_model->ExternalLinks = array();
            $this->_model->ExternalLinks = array_merge($this->_model->ExternalLinks, $loadedModel->ExternalLinks);
        } catch (CellsException $exception) {
            throw $exception;
        } catch (\Throwable $exception) {
            throw new WorkbookLoadException('Failed to load XLSX workbook.', $exception);
        }
    }
}
