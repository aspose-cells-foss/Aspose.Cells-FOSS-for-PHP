<?php

namespace Aspose\Cells_FOSS\Testing;

use Aspose\Cells_FOSS\ImageType;
use Aspose\Cells_FOSS\Workbook;

final class PictureScenarioFactory
{
    public static function LoadJpeg()
    {
        $path = self::ResolveInputFile('pay.jpg');
        $data = @file_get_contents($path);
        if ($data === false) {
            throw new \RuntimeException('Could not open Input/pay.jpg.');
        }

        return $data;
    }

    public static function LoadPng()
    {
        $path = self::ResolveInputFile('screen.png');
        $data = @file_get_contents($path);
        if ($data === false) {
            throw new \RuntimeException('Could not open Input/screen.png.');
        }

        return $data;
    }

    public static function CreatePictureWorkbook()
    {
        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Images';

        $sheet->Cells['A1']->PutValue('Name');
        $sheet->Cells['B1']->PutValue('Value');
        $sheet->Cells['A2']->PutValue('Alpha');
        $sheet->Cells['B2']->PutValue(42);

        $jpegIndex = $sheet->Pictures->Add(0, 0, 8, 5, self::LoadJpeg());
        $sheet->Pictures[$jpegIndex]->Name = 'PayImage';

        $pngIndex = $sheet->Pictures->Add(10, 0, 16, 5, self::LoadPng());
        $sheet->Pictures[$pngIndex]->Name = 'ScreenImage';

        return $workbook;
    }

    public static function AssertPictureWorkbook(Workbook $workbook)
    {
        $sheet = $workbook->Worksheets[0];
        AssertEx::Equal(2, $sheet->Pictures->Count);

        $jpeg = $sheet->Pictures[0];
        AssertEx::Equal('PayImage', $jpeg->Name);
        AssertEx::Equal(0, $jpeg->UpperLeftRow);
        AssertEx::Equal(0, $jpeg->UpperLeftColumn);
        AssertEx::Equal(8, $jpeg->LowerRightRow);
        AssertEx::Equal(5, $jpeg->LowerRightColumn);
        AssertEx::Equal(ImageType::Jpeg, $jpeg->ImageType);

        $png = $sheet->Pictures[1];
        AssertEx::Equal('ScreenImage', $png->Name);
        AssertEx::Equal(10, $png->UpperLeftRow);
        AssertEx::Equal(0, $png->UpperLeftColumn);
        AssertEx::Equal(16, $png->LowerRightRow);
        AssertEx::Equal(5, $png->LowerRightColumn);
        AssertEx::Equal(ImageType::Png, $png->ImageType);
    }

    private static function ResolveInputFile($fileName)
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== '') {
            if (is_file($directory . DIRECTORY_SEPARATOR . 'composer.json')) {
                return $directory . DIRECTORY_SEPARATOR . 'Input' . DIRECTORY_SEPARATOR . $fileName;
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }

            $directory = $parent;
        }

        throw new \RuntimeException('Could not locate repository root to resolve Input/' . $fileName . '.');
    }
}
