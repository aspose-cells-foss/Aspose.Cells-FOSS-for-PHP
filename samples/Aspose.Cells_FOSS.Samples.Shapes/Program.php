<?php

namespace Aspose\Cells_FOSS\Samples\Shapes;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\AutoShapeType;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $outputPath = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'shapes-sample.xlsx';

        $workbook = new Workbook();
        $sheet = $workbook->Worksheets[0];
        $sheet->Name = 'Shapes';

        $sheet->Cells['A1']->PutValue('Shape');
        $sheet->Cells['B1']->PutValue('Description');

        $rectIndex = $sheet->Shapes->Add(0, 2, 2, 3, AutoShapeType::Rectangle);
        $sheet->Shapes[$rectIndex]->Name = 'Rectangle Box';

        $roundedRectIndex = $sheet->Shapes->Add(0, 4, 2, 5, AutoShapeType::RoundedRectangle);
        $sheet->Shapes[$roundedRectIndex]->Name = 'Rounded Rectangle';

        $ellipseIndex = $sheet->Shapes->Add(0, 6, 2, 7, AutoShapeType::Ellipse);
        $sheet->Shapes[$ellipseIndex]->Name = 'Oval Shape';

        $triangleIndex = $sheet->Shapes->Add(3, 2, 5, 3, AutoShapeType::Triangle);
        $sheet->Shapes[$triangleIndex]->Name = 'Triangle';

        $rightTriangleIndex = $sheet->Shapes->Add(3, 4, 5, 5, AutoShapeType::RightTriangle);
        $sheet->Shapes[$rightTriangleIndex]->Name = 'Right Triangle';

        $diamondIndex = $sheet->Shapes->Add(3, 6, 5, 7, AutoShapeType::Diamond);
        $sheet->Shapes[$diamondIndex]->Name = 'Diamond';

        $pentagonIndex = $sheet->Shapes->Add(6, 2, 8, 3, AutoShapeType::Pentagon);
        $sheet->Shapes[$pentagonIndex]->Name = 'Pentagon';

        $hexagonIndex = $sheet->Shapes->Add(6, 4, 8, 5, AutoShapeType::Hexagon);
        $sheet->Shapes[$hexagonIndex]->Name = 'Hexagon';

        $octagonIndex = $sheet->Shapes->Add(6, 6, 8, 7, AutoShapeType::Octagon);
        $sheet->Shapes[$octagonIndex]->Name = 'Octagon';

        $star5Index = $sheet->Shapes->Add(9, 2, 11, 3, AutoShapeType::Star5Point);
        $sheet->Shapes[$star5Index]->Name = '5-Point Star';

        $star8Index = $sheet->Shapes->Add(9, 4, 11, 5, AutoShapeType::Star8Point);
        $sheet->Shapes[$star8Index]->Name = '8-Point Star';

        $star12Index = $sheet->Shapes->Add(9, 6, 11, 7, AutoShapeType::Star12Point);
        $sheet->Shapes[$star12Index]->Name = '12-Point Star';

        $rightArrowIndex = $sheet->Shapes->Add(12, 2, 14, 3, AutoShapeType::RightArrow);
        $sheet->Shapes[$rightArrowIndex]->Name = 'Right Arrow';

        $leftArrowIndex = $sheet->Shapes->Add(12, 4, 14, 5, AutoShapeType::LeftArrow);
        $sheet->Shapes[$leftArrowIndex]->Name = 'Left Arrow';

        $upDownArrowIndex = $sheet->Shapes->Add(12, 6, 14, 7, AutoShapeType::UpDownArrow);
        $sheet->Shapes[$upDownArrowIndex]->Name = 'Up-Down Arrow';

        $heartIndex = $sheet->Shapes->Add(15, 2, 17, 3, AutoShapeType::Heart);
        $sheet->Shapes[$heartIndex]->Name = 'Heart';

        $lightningIndex = $sheet->Shapes->Add(15, 4, 17, 5, AutoShapeType::Lightning);
        $sheet->Shapes[$lightningIndex]->Name = 'Lightning Bolt';

        $cloudIndex = $sheet->Shapes->Add(15, 6, 17, 7, AutoShapeType::Cloud);
        $sheet->Shapes[$cloudIndex]->Name = 'Cloud';

        $sunIndex = $sheet->Shapes->Add(18, 2, 20, 3, AutoShapeType::Sun);
        $sheet->Shapes[$sunIndex]->Name = 'Sun';

        $moonIndex = $sheet->Shapes->Add(18, 4, 20, 5, AutoShapeType::Moon);
        $sheet->Shapes[$moonIndex]->Name = 'Moon';

        $plusIndex = $sheet->Shapes->Add(18, 6, 20, 7, AutoShapeType::Plus);
        $sheet->Shapes[$plusIndex]->Name = 'Plus Sign';

        $cubeIndex = $sheet->Shapes->Add(21, 2, 23, 3, AutoShapeType::Cube);
        $sheet->Shapes[$cubeIndex]->Name = 'Cube';

        $cylinderIndex = $sheet->Shapes->Add(21, 4, 23, 5, AutoShapeType::Cylinder);
        $sheet->Shapes[$cylinderIndex]->Name = 'Cylinder';

        $mathPlusIndex = $sheet->Shapes->Add(21, 6, 23, 7, AutoShapeType::MathPlus);
        $sheet->Shapes[$mathPlusIndex]->Name = 'Math Plus';

        $workbook->Save($outputPath);

        $loaded = new Workbook($outputPath);
        $loadedSheet = $loaded->Worksheets['Shapes'];

        echo 'Saved: ' . $outputPath . PHP_EOL;
        echo 'Shape count: ' . $loadedSheet->Shapes->Count . PHP_EOL;
        echo PHP_EOL;
        echo 'Basic Shapes:' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[0]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[0]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[1]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[1]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[2]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[2]->AutoShapeType) . ')' . PHP_EOL;
        echo PHP_EOL;
        echo 'Triangles & Polygons:' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[3]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[3]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[4]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[4]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[5]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[5]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[6]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[6]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[7]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[7]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[8]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[8]->AutoShapeType) . ')' . PHP_EOL;
        echo PHP_EOL;
        echo 'Stars:' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[9]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[9]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[10]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[10]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[11]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[11]->AutoShapeType) . ')' . PHP_EOL;
        echo PHP_EOL;
        echo 'Arrows:' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[12]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[12]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[13]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[13]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[14]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[14]->AutoShapeType) . ')' . PHP_EOL;
        echo PHP_EOL;
        echo 'Symbols & Icons:' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[15]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[15]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[16]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[16]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[17]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[17]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[18]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[18]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[19]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[19]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[20]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[20]->AutoShapeType) . ')' . PHP_EOL;
        echo PHP_EOL;
        echo '3D & Math:' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[21]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[21]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[22]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[22]->AutoShapeType) . ')' . PHP_EOL;
        echo '  ' . $loadedSheet->Shapes[23]->Name . ' (' . self::AutoShapeTypeName($loadedSheet->Shapes[23]->AutoShapeType) . ')' . PHP_EOL;
    }

    private static function AutoShapeTypeName($value)
    {
        switch ($value) {
            case AutoShapeType::Unknown:
                return 'Unknown';
            case AutoShapeType::Rectangle:
                return 'Rectangle';
            case AutoShapeType::RoundedRectangle:
                return 'RoundedRectangle';
            case AutoShapeType::Ellipse:
                return 'Ellipse';
            case AutoShapeType::Triangle:
                return 'Triangle';
            case AutoShapeType::RightTriangle:
                return 'RightTriangle';
            case AutoShapeType::Diamond:
                return 'Diamond';
            case AutoShapeType::Pentagon:
                return 'Pentagon';
            case AutoShapeType::Hexagon:
                return 'Hexagon';
            case AutoShapeType::Octagon:
                return 'Octagon';
            case AutoShapeType::Plus:
                return 'Plus';
            case AutoShapeType::Cube:
                return 'Cube';
            case AutoShapeType::Cylinder:
                return 'Cylinder';
            case AutoShapeType::Heart:
                return 'Heart';
            case AutoShapeType::Lightning:
                return 'Lightning';
            case AutoShapeType::Sun:
                return 'Sun';
            case AutoShapeType::Moon:
                return 'Moon';
            case AutoShapeType::Cloud:
                return 'Cloud';
            case AutoShapeType::RightArrow:
                return 'RightArrow';
            case AutoShapeType::LeftArrow:
                return 'LeftArrow';
            case AutoShapeType::UpArrow:
                return 'UpArrow';
            case AutoShapeType::DownArrow:
                return 'DownArrow';
            case AutoShapeType::LeftRightArrow:
                return 'LeftRightArrow';
            case AutoShapeType::UpDownArrow:
                return 'UpDownArrow';
            case AutoShapeType::Star4Point:
                return 'Star4Point';
            case AutoShapeType::Star5Point:
                return 'Star5Point';
            case AutoShapeType::Star6Point:
                return 'Star6Point';
            case AutoShapeType::Star7Point:
                return 'Star7Point';
            case AutoShapeType::Star8Point:
                return 'Star8Point';
            case AutoShapeType::Star10Point:
                return 'Star10Point';
            case AutoShapeType::Star12Point:
                return 'Star12Point';
            case AutoShapeType::Star16Point:
                return 'Star16Point';
            case AutoShapeType::Star24Point:
                return 'Star24Point';
            case AutoShapeType::Star32Point:
                return 'Star32Point';
            case AutoShapeType::TextBox:
                return 'TextBox';
            case AutoShapeType::MathPlus:
                return 'MathPlus';
            case AutoShapeType::StraightConnector:
                return 'StraightConnector';
            case AutoShapeType::BentConnector:
                return 'BentConnector';
            case AutoShapeType::CurvedConnector:
                return 'CurvedConnector';
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
