# Aspose.Cells FOSS for PHP

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](License/LICENSE.txt) [![PHP](https://img.shields.io/badge/php-%3E%3D8.1-777bb4.svg)](https://www.php.net/)

Aspose.Cells FOSS for PHP is an MIT-licensed spreadsheet library for PHP. It provides an Aspose.Cells-style API for creating, loading, editing, and saving Excel `.xlsx` workbooks, based on a faithful port of the C# `aspose-cells-foss-for-net` sources.

The package requires PHP 8.1 or later with the `zip`, `dom`, `simplexml`, and `mbstring` extensions. It runs on Windows, Linux, and macOS, and does not require Microsoft Excel or any native Office installation. The Composer package name is `aspose/cells-foss-for-php`; the public namespace is `Aspose\Cells_FOSS`.

## Navigation

- [Key Capabilities](#key-capabilities)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Examples](#examples)
- [Documentation](#documentation)
- [Scope and Limitations](#scope-and-limitations)
- [Development and Testing](#development-and-testing)
- [License](#license)

## Key Capabilities

- Create, load, modify, and save Excel `.xlsx` workbooks with `Workbook`, `Worksheet`, `Cells`, and `Cell`.
- Load malformed packages with recovery-oriented `LoadOptions` and inspect `LoadDiagnostics`.
- Read and write scalar values, formulas, dates, and number formats.
- Apply styles including fonts, fills, borders, alignment, wrapping, and locking.
- Configure validations, conditional formatting, hyperlinks, defined names, filters, tables, comments, pictures, shapes, charts, and page setup.

## Installation

```bash
composer require aspose/cells-foss-for-php
```

The library uses PSR-4 autoloading under the `Aspose\Cells_FOSS\` namespace. When Composer is not available, a bootstrap autoloader is included at `autoload.php`:

```php
require __DIR__ . '/autoload.php';
```

## Quick Start

```php
<?php

require __DIR__ . '/vendor/autoload.php';

use Aspose\Cells_FOSS\Color;
use Aspose\Cells_FOSS\FillPattern;
use Aspose\Cells_FOSS\Workbook;

$workbook = new Workbook();
$sheet = $workbook->Worksheets[0];

$sheet->Name = 'Products';
$sheet->Cells['A1']->PutValue('Product');
$sheet->Cells['B1']->PutValue('Price');
$sheet->Cells['A2']->PutValue('Apple');
$sheet->Cells['B2']->PutValue(2.99);
$sheet->Cells['A3']->PutValue('Orange');
$sheet->Cells['B3']->PutValue(1.99);
$sheet->Cells['B4']->Formula = '=SUM(B2:B3)';

$headerStyle = $sheet->Cells['A1']->GetStyle();
$headerStyle->Font->IsBold = true;
$headerStyle->Font->Color = Color::FromArgb(255, 255, 255, 255);
$headerStyle->Pattern = FillPattern::Solid;
$headerStyle->ForegroundColor = Color::FromArgb(255, 34, 120, 212);

$sheet->Cells['A1']->SetStyle($headerStyle);
$sheet->Cells['B1']->SetStyle($headerStyle);

$workbook->Save('products.xlsx');
```

Load an existing workbook with repair-friendly options:

```php
<?php

require __DIR__ . '/vendor/autoload.php';

use Aspose\Cells_FOSS\LoadOptions;
use Aspose\Cells_FOSS\Workbook;

$loadOptions = new LoadOptions();
$loadOptions->TryRepairPackage = true;
$loadOptions->TryRepairXml = true;
$loadOptions->StrictMode = false;

$workbook = new Workbook('input.xlsx', $loadOptions);

if ($workbook->LoadDiagnostics->HasDataLossRisk) {
    echo 'Potential data loss risk detected during load.' . PHP_EOL;
}

$workbook->Worksheets[0]->Cells['A1']->PutValue('Updated');
$workbook->Save('updated.xlsx');
```

## Examples

Runnable samples live under [`samples/`](samples/README.md). Feature areas currently covered include:

- basic workbook and cell operations
- load and repair behavior
- styles
- worksheet settings
- validations
- conditional formatting
- hyperlinks and defined names
- page setup
- comments
- charts
- list objects
- pictures
- document properties
- shapes

Run one sample from the repository root:

```bash
php samples/Aspose.Cells_FOSS.Samples.Basic/Program.php
```

Generated workbooks are written to `output/samples` under the repository root.

## Documentation

- [Samples overview](samples/README.md)
- [C# to PHP porting conventions](PORTING_CONVENTIONS.md)
- [Bootstrap autoloader](autoload.php)

## Scope and Limitations

- Only `.xlsx` is supported as the spreadsheet file format for load and save.
- The library stores formulas but does not include a formula calculation engine.
- There is no PDF, HTML, or image export API.
- Macros, VBA, and pivot tables are outside the supported public API surface.
- Broader format coverage and richer spreadsheet features belong to the commercial Aspose.Cells product line.

## Development and Testing

Lint a single source file:

```bash
php -l src/Workbook.php
```

Run the automated test programs from the repository root:

```bash
powershell -File run-cell-data-tests.ps1
powershell -File run-compare-openxml.ps1
powershell -File run-page-setup-openxml-tests.ps1
powershell -File run-worksheet-settings-openxml-tests.ps1
```

The repository also includes focused test programs for workbook compatibility, page setup, conditional formatting, data validation, chart Open XML, supported features, and golden files under `tests/`.

## License

This project is licensed under the [MIT License](License/LICENSE.txt).
