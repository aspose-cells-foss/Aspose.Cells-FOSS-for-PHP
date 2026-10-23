<?php

namespace Aspose\Cells_FOSS\Samples\HyperlinksAndNames;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\TargetModeType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'hyperlinks-and-names-sample.xlsx';

        $workbook = new Workbook();
        $linksSheet = $workbook->Worksheets[0];
        $linksSheet->Name = 'Links';
        $scopedSheetIndex = $workbook->Worksheets->Add('Scoped Sheet');
        $scopedSheet = $workbook->Worksheets[$scopedSheetIndex];
        $scopedSheet->Cells['B2']->PutValue(5);

        $linksSheet->Cells['A1']->PutValue('Docs');
        $externalLink = $linksSheet->Hyperlinks[$linksSheet->Hyperlinks->Add('A1', 1, 1, 'https://example.com/docs?q=1')];
        $externalLink->TextToDisplay = 'Docs';
        $externalLink->ScreenTip = 'External docs';

        $linksSheet->Cells['B2']->PutValue('Jump');
        $internalLink = $linksSheet->Hyperlinks[$linksSheet->Hyperlinks->Add('B2', 1, 1, "'Scoped Sheet'!B2")];
        $internalLink->TextToDisplay = 'Jump';
        $internalLink->ScreenTip = 'Jump to scoped sheet';

        $mailLink = $linksSheet->Hyperlinks[$linksSheet->Hyperlinks->Add('C4', 'D5', 'mailto:test@example.com', 'Mail', 'Send mail')];
        $mailLink->ScreenTip = 'Send mail';

        $globalName = $workbook->DefinedNames[$workbook->DefinedNames->Add('GlobalRange', "='Links'!\$A\$1:\$D\$5")];
        $globalName->Hidden = true;
        $globalName->Comment = 'Primary sample range';

        $localName = $workbook->DefinedNames[$workbook->DefinedNames->Add('ScopedCell', "'Scoped Sheet'!\$B\$2", $scopedSheetIndex)];
        $localName->Comment = 'Sheet-scoped name';

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedLinks = $loaded->Worksheets['Links'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Hyperlinks: ' . $loadedLinks->Hyperlinks->Count . PHP_EOL;
        echo 'First hyperlink: ' . $loadedLinks->Hyperlinks[0]->Address . ' / ' . self::TargetModeTypeName($loadedLinks->Hyperlinks[0]->LinkType) . PHP_EOL;
        echo 'Second hyperlink: ' . $loadedLinks->Hyperlinks[1]->Address . ' / ' . self::TargetModeTypeName($loadedLinks->Hyperlinks[1]->LinkType) . PHP_EOL;
        echo 'Defined names: ' . $loaded->DefinedNames->Count . PHP_EOL;
        echo 'Global name formula: ' . $loaded->DefinedNames[0]->Formula . PHP_EOL;
        echo 'Local name scope: ' . self::GetSheetIndexOrDefault($loaded->DefinedNames[1]->LocalSheetIndex) . PHP_EOL;
    }

    private static function GetSheetIndexOrDefault($localSheetIndex)
    {
        if ($localSheetIndex === null) {
            return -1;
        }

        return $localSheetIndex;
    }

    private static function TargetModeTypeName($value)
    {
        switch ($value) {
            case TargetModeType::External:
                return 'External';
            case TargetModeType::FilePath:
                return 'FilePath';
            case TargetModeType::Email:
                return 'Email';
            case TargetModeType::CellReference:
                return 'CellReference';
            default:
                return (string)$value;
        }
    }

    private static function ResolveOutputDirectory()
    {
        $directory = self::ResolveRepositoryRoot() . DIRECTORY_SEPARATOR . 'output' . DIRECTORY_SEPARATOR . 'samples';
        if (!is_dir($directory)) {
            mkdir($directory, 0777, true);
        }

        return $directory;
    }

    private static function ResolveRepositoryRoot()
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
