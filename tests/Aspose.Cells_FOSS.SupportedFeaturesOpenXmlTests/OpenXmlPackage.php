<?php

namespace Aspose\Cells_FOSS\SupportedFeaturesOpenXmlTests;

use Aspose\Cells_FOSS\Internal\Xml\XDocument;

final class OpenXmlPackage
{
    const RelationshipNs = '{http://schemas.openxmlformats.org/package/2006/relationships}';

    private $filePath;
    private $archive;
    private $entries = array();

    public function __construct($filePath)
    {
        $this->filePath = $filePath;
        $this->archive = new \ZipArchive();
        $result = $this->archive->open($filePath);
        if ($result !== true) {
            throw new \RuntimeException('Could not open Open XML package: ' . $filePath);
        }

        for ($index = 0; $index < $this->archive->numFiles; $index++) {
            $name = $this->archive->getNameIndex($index);
            if ($name !== false) {
                $this->entries[ltrim($name, '/')] = $index;
            }
        }
    }

    public function GetFilePath()
    {
        return $this->filePath;
    }

    public function Exists($partName)
    {
        $key = ltrim($partName, '/');
        return array_key_exists($key, $this->entries);
    }

    public function ReadText($partName)
    {
        $key = ltrim($partName, '/');
        if (!array_key_exists($key, $this->entries)) {
            throw new \RuntimeException('Package part not found: ' . $partName);
        }

        $content = $this->archive->getFromIndex($this->entries[$key]);
        if ($content === false) {
            throw new \RuntimeException('Could not read package part: ' . $partName);
        }

        return $content;
    }

    public function ReadXml($partName)
    {
        return XDocument::Load($this->ReadText($partName));
    }

    public function ResolvePart($sourcePart, $relationshipTarget)
    {
        if ($relationshipTarget === null) {
            return null;
        }

        $target = str_replace('\\', '/', trim($relationshipTarget));
        if ($target === '') {
            return null;
        }

        if (strpos($target, '/') === 0) {
            return trim($this->Normalize(ltrim($target, '/')), '/');
        }

        $baseDirectory = '';
        $slashIndex = strrpos($sourcePart, '/');
        if ($slashIndex !== false) {
            $baseDirectory = substr($sourcePart, 0, $slashIndex + 1);
        }

        return trim($this->Normalize($baseDirectory . $target), '/');
    }

    private function Normalize($path)
    {
        $segments = explode('/', $path);
        $stack = array();
        foreach ($segments as $segment) {
            if ($segment === '' || $segment === '.') {
                continue;
            }

            if ($segment === '..') {
                if (count($stack) > 0) {
                    array_pop($stack);
                }
                continue;
            }

            $stack[] = $segment;
        }

        return implode('/', $stack);
    }

    public function ReadRelationships($sourcePart)
    {
        $relationships = array();
        $slashIndex = strrpos($sourcePart, '/');
        $directory = $slashIndex === false ? '' : substr($sourcePart, 0, $slashIndex);
        $fileName = $slashIndex === false ? $sourcePart : substr($sourcePart, $slashIndex + 1);
        $relsPath = ($directory === '' ? '' : $directory . '/') . '_rels/' . $fileName . '.rels';

        if (!$this->Exists($relsPath)) {
            return $relationships;
        }

        $document = $this->ReadXml($relsPath);
        $relationshipsElement = $document->Root;
        if ($relationshipsElement === null) {
            return $relationships;
        }

        foreach ($relationshipsElement->Elements(self::RelationshipNs . 'Relationship') as $relationship) {
            $id = ComparisonValueHelpers::AttributeText($relationship, 'Id');
            if ($id === null) {
                continue;
            }

            $type = ComparisonValueHelpers::AttributeText($relationship, 'Type');
            $target = ComparisonValueHelpers::AttributeText($relationship, 'Target');
            $targetMode = ComparisonValueHelpers::AttributeText($relationship, 'TargetMode');
            $isExternal = $targetMode !== null && strcasecmp($targetMode, 'External') === 0;

            $relationships[$id] = array(
                'Type' => $type === null ? '' : $type,
                'Target' => $target === null ? '' : $target,
                'TargetMode' => $targetMode,
                'IsExternal' => $isExternal,
                'ResolvedTarget' => $isExternal ? $target : $this->ResolvePart($sourcePart, $target),
            );
        }

        return $relationships;
    }
}
