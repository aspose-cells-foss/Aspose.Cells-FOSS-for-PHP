<?php

namespace Aspose\Cells_FOSS\Samples\Comments;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'comments-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Comments';

        $sheet->Cells['A1']->PutValue('Task');
        $sheet->Cells['B1']->PutValue('Status');
        $sheet->Cells['A2']->PutValue('Design UI');
        $sheet->Cells['B2']->PutValue('In Progress');
        $sheet->Cells['A3']->PutValue('Implement API');
        $sheet->Cells['B3']->PutValue('Pending');
        $sheet->Cells['A4']->PutValue('Write Tests');
        $sheet->Cells['B4']->PutValue('Completed');

        $comment2 = $sheet->Comments->Add(1, 0);
        $comment2->Author = 'John';
        $comment2->Note = 'Need to finalize the color scheme by Friday';
        $comment2->IsVisible = false;
        $comment2->Width = 200;
        $comment2->Height = 100;

        $comment3 = $sheet->Comments->Add('A3');
        $comment3->Author = 'Sarah';
        $comment3->Note = 'Waiting for database schema approval';
        $comment3->IsVisible = true;
        $comment3->Width = 180;
        $comment3->Height = 90;

        $comment4 = $sheet->Comments->Add('B4');
        $comment4->Author = 'Mike';
        $comment4->Note = 'All unit tests passing. Ready for review.';
        $comment4->IsVisible = false;

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Comments'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Comment count: ' . $loadedSheet->Comments->Count . PHP_EOL;
        echo 'A2 comment by ' . $loadedSheet->Comments['A2']->Author . ': ' . $loadedSheet->Comments['A2']->Note . PHP_EOL;
        echo 'A3 comment visible: ' . self::FormatBool($loadedSheet->Comments['A3']->IsVisible) . PHP_EOL;
        echo 'B4 comment: ' . $loadedSheet->Comments['B4']->Note . PHP_EOL;
    }

    private static function FormatBool($value)
    {
        return $value ? 'True' : 'False';
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
