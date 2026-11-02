<?php

namespace Aspose\Cells_FOSS\Samples\PdfConversion;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\PdfSaveOptions;
use Aspose\Cells_FOSS\Workbook;

final class Program
{
    public static function Main()
    {
        $repositoryRoot = self::ResolveRepositoryRoot();
        $evaluationDate = null;
        $inputArguments = array();
        $arguments = array_slice($_SERVER['argv'], 1);

        for ($index = 0; $index < count($arguments); $index++) {
            if ($arguments[$index] === '--conditional-date') {
                $next = $index + 1 < count($arguments) ? $arguments[$index + 1] : null;
                $parsedDate = \DateTime::createFromFormat('Y-m-d', $next !== null ? $next : '');
                if ($parsedDate === false) {
                    throw new \InvalidArgumentException('--conditional-date requires a yyyy-MM-dd date.');
                }

                $evaluationDate = $parsedDate;
                $index++;
                continue;
            }

            $inputArguments[] = $arguments[$index];
        }

        $inputPaths = self::ResolveInputPaths($repositoryRoot, $inputArguments);
        $outputDirectory = self::ResolveOutputDirectory() . DIRECTORY_SEPARATOR . 'pdf';
        if (!is_dir($outputDirectory)) {
            mkdir($outputDirectory, 0777, true);
        }

        foreach ($inputPaths as $inputPath) {
            $outputPath = $outputDirectory . DIRECTORY_SEPARATOR . pathinfo($inputPath, PATHINFO_FILENAME) . '.pdf';

            $workbook = new Workbook($inputPath);
            $options = new PdfSaveOptions();
            $options->ConditionalFormattingEvaluationDate = $evaluationDate;
            $workbook->Save($outputPath, $options);

            echo 'Input: ' . $inputPath . PHP_EOL;
            echo 'Output: ' . $outputPath . PHP_EOL;
        }
    }

    private static function ResolveInputPaths($repositoryRoot, $inputArguments)
    {
        if (count($inputArguments) > 0) {
            return self::ResolveExplicitPaths($repositoryRoot, $inputArguments);
        }

        return self::ResolveAllInputPaths($repositoryRoot);
    }

    private static function ResolveExplicitPaths($repositoryRoot, $inputArguments)
    {
        $results = array();
        foreach ($inputArguments as $inputArgument) {
            if (trim($inputArgument) === '') {
                continue;
            }

            $results[] = self::ResolveExplicitPath($repositoryRoot, $inputArgument);
        }

        if (count($results) === 0) {
            throw new \InvalidArgumentException('At least one non-empty .xlsx path must be provided.');
        }

        return $results;
    }

    private static function ResolveAllInputPaths($repositoryRoot)
    {
        $inputDirectory = $repositoryRoot . DIRECTORY_SEPARATOR . 'samples' . DIRECTORY_SEPARATOR . 'files';
        $results = array();
        foreach (glob($inputDirectory . DIRECTORY_SEPARATOR . '*.xlsx') as $file) {
            $results[] = $file;
        }

        usort($results, 'strcasecmp');

        if (count($results) === 0) {
            throw new \RuntimeException('Could not find any .xlsx files under the samples/files directory.');
        }

        return $results;
    }

    private static function ResolveExplicitPath($repositoryRoot, $inputArgument)
    {
        $candidatePath = $inputArgument;
        if (!self::IsRooted($candidatePath)) {
            $candidatePath = $repositoryRoot . DIRECTORY_SEPARATOR . $candidatePath;
            if (!is_file($candidatePath)) {
                $candidatePath = $repositoryRoot . DIRECTORY_SEPARATOR . 'samples' . DIRECTORY_SEPARATOR . 'files'
                    . DIRECTORY_SEPARATOR . $inputArgument;
            }
        }

        if (!is_file($candidatePath)) {
            throw new \RuntimeException('The specified XLSX file does not exist: ' . $candidatePath);
        }

        if (strtolower(pathinfo($candidatePath, PATHINFO_EXTENSION)) !== 'xlsx') {
            throw new \InvalidArgumentException('The input file must have an .xlsx extension.');
        }

        return realpath($candidatePath);
    }

    private static function IsRooted($path)
    {
        return str_starts_with($path, DIRECTORY_SEPARATOR) || preg_match('/^[A-Za-z]:/', $path) === 1;
    }

    private static function ResolveOutputDirectory()
    {
        $directory = __DIR__;
        while ($directory !== null && $directory !== '') {
            if (is_file($directory . DIRECTORY_SEPARATOR . 'composer.json')) {
                return $directory . DIRECTORY_SEPARATOR . 'output' . DIRECTORY_SEPARATOR . 'samples';
            }

            $parent = dirname($directory);
            if ($parent === $directory) {
                break;
            }

            $directory = $parent;
        }

        return getcwd();
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
