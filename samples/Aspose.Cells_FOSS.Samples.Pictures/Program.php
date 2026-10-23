<?php

namespace Aspose\Cells_FOSS\Samples\Pictures;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\ImageType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'pictures-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Pictures';

        $sheet->Cells['A1']->PutValue('Product');
        $sheet->Cells['B1']->PutValue('Description');
        $sheet->Cells['A2']->PutValue('Item 1');
        $sheet->Cells['B2']->PutValue('Sample product with logo');
        $sheet->Cells['A4']->PutValue('Item 2');
        $sheet->Cells['B4']->PutValue('Another product');

        $sampleImageData = self::CreateSamplePngImage();

        $picture1Index = $sheet->Pictures->Add(1, 2, 3, 3, $sampleImageData);
        $sheet->Pictures[$picture1Index]->Name = 'Product Logo 1';

        $picture2Index = $sheet->Pictures->Add(3, 2, 5, 3, $sampleImageData);
        $picture2 = $sheet->Pictures[$picture2Index];
        $picture2->Name = 'Product Logo 2';
        $picture2->UpperLeftColumn = 4;
        $picture2->LowerRightColumn = 5;

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Pictures'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Picture count: ' . $loadedSheet->Pictures->Count . PHP_EOL;
        echo 'First picture: ' . $loadedSheet->Pictures[0]->Name . ' (Type: ' . self::ImageTypeName($loadedSheet->Pictures[0]->ImageType) . ')' . PHP_EOL;
        echo 'First picture anchor: R' . $loadedSheet->Pictures[0]->UpperLeftRow . 'C' . $loadedSheet->Pictures[0]->UpperLeftColumn . ' to R' . $loadedSheet->Pictures[0]->LowerRightRow . 'C' . $loadedSheet->Pictures[0]->LowerRightColumn . PHP_EOL;
        echo 'Second picture: ' . $loadedSheet->Pictures[1]->Name . ' (Type: ' . self::ImageTypeName($loadedSheet->Pictures[1]->ImageType) . ')' . PHP_EOL;
    }

    private static function CreateSamplePngImage()
    {
        $image = '';
        $image .= "\x89\x50\x4E\x47\x0D\x0A\x1A\x0A";
        $image .= self::WriteChunk('IHDR', "\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00");
        $image .= self::WriteChunk('IDAT', "\x78\x9C\x62\x00\x02\x00\x00\x05\x00\x01\x0D");
        $image .= self::WriteChunk('IEND', '');

        return $image;
    }

    private static function WriteChunk($type, $data)
    {
        return pack('N', strlen($data)) . $type . $data . pack('N', self::CalculateCrc($type . $data));
    }

    private static function CalculateCrc($buffer)
    {
        $crc = 0xFFFFFFFF;
        $length = strlen($buffer);
        for ($i = 0; $i < $length; $i++) {
            $crc ^= ord($buffer[$i]);
            for ($j = 0; $j < 8; $j++) {
                if (($crc & 1) !== 0) {
                    $crc = ($crc >> 1) ^ 0xEDB88320;
                } else {
                    $crc >>= 1;
                }
            }
        }

        return (~$crc) & 0xFFFFFFFF;
    }

    private static function ImageTypeName($value)
    {
        switch ($value) {
            case ImageType::Unknown:
                return 'Unknown';
            case ImageType::Jpeg:
                return 'Jpeg';
            case ImageType::Png:
                return 'Png';
            case ImageType::Gif:
                return 'Gif';
            case ImageType::Bmp:
                return 'Bmp';
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
