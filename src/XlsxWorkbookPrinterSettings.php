<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\PageOrientation;
use Aspose\Cells_FOSS\Core\PageSetupModel;
use Aspose\Cells_FOSS\Internal\Io\ZipArchive;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookPrinterSettings
{
    private function __construct()
    {
    }

    public static function Load(PageSetupModel $setup, XElement $root, ZipArchive $archive, $worksheetUri, $diagnostics, $options, $sheetName)
    {
        $main = XlsxWorkbookSerializerCommon::MainNs;
        $rel = XlsxWorkbookSerializerCommon::RelationshipNs;
        $pageSetupElement = $root->Element($main . 'pageSetup');
        $id = $pageSetupElement === null ? null : (string) $pageSetupElement->Attribute($rel . 'id');
        if ($id === null || $id === ''
            || ($setup->PaperSize !== 0 && $setup->Orientation !== PageOrientation::Default)) {
            return;
        }

        $slash = strrpos($worksheetUri, '/');
        $relsUri = substr($worksheetUri, 0, $slash + 1) . '_rels/' . substr($worksheetUri, $slash + 1) . '.rels';
        $relationships = XlsxWorkbookArchiveHelpers::LoadRelationships($archive, $relsUri, $worksheetUri);
        if (!array_key_exists($id, $relationships)) {
            self::Warn($diagnostics, $options, $sheetName);
            return;
        }

        $target = $relationships[$id];
        $entry = XlsxWorkbookArchiveHelpers::GetEntry($archive, $target);
        $content = $entry === null ? '' : $entry->getContent();
        if ($entry === null || strlen($content) < 220) {
            self::Warn($diagnostics, $options, $sheetName);
            return;
        }

        $header = substr($content, 0, 220);
        if (!self::TryApply($header, strlen($content), $setup)) {
            self::Warn($diagnostics, $options, $sheetName);
        }
    }

    public static function TryApply($header, $length, PageSetupModel $setup)
    {
        if ($header === null || strlen($header) < 220) {
            return false;
        }

        $size = self::ReadUInt16($header, 68);
        $extra = self::ReadUInt16($header, 70);
        if ($size < 220 || $size + $extra > $length) {
            return false;
        }

        // dmFields marks which fields are meaningful. Never interpret driver-private bytes.
        $fields = self::ReadUInt16($header, 72);
        $paper = self::ReadUInt16($header, 78);
        if (($fields & 2) !== 0 && $paper > 0 && $paper < 256 && $setup->PaperSize === 0) {
            $setup->PaperSize = $paper;
        }

        $orientation = self::ReadUInt16($header, 76);
        if (($fields & 1) !== 0 && $setup->Orientation === PageOrientation::Default) {
            if ($orientation === 1) {
                $setup->Orientation = PageOrientation::Portrait;
            }

            if ($orientation === 2) {
                $setup->Orientation = PageOrientation::Landscape;
            }
        }

        return true;
    }

    private static function ReadUInt16($data, $offset)
    {
        return ord($data[$offset]) | ord($data[$offset + 1]) << 8;
    }

    private static function Warn($diagnostics, $options, $sheetName)
    {
        $issue = new LoadIssue('PS-L001', DiagnosticSeverity::Warning,
            'Saved printer settings are missing or invalid; explicit page settings and default paper size were retained.', false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
    }
}
