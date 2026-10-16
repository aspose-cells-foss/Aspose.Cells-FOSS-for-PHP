<?php

namespace Aspose\Cells_FOSS;

final class LoadDiagnostics
{
    private $_issues = array();

    public function __get($propertyName)
    {
        if ($propertyName === 'Issues') {
            return $this->_issues;
        }
        if ($propertyName === 'HasRepairs') {
            foreach ($this->_issues as $issue) {
                if ($issue->RepairApplied) {
                    return true;
                }
            }

            return false;
        }
        if ($propertyName === 'HasDataLossRisk') {
            foreach ($this->_issues as $issue) {
                if ($issue->DataLossRisk) {
                    return true;
                }
            }

            return false;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Issues') {
            throw new \RuntimeException('Issues is read-only.');
        }
        if ($propertyName === 'HasRepairs') {
            throw new \RuntimeException('HasRepairs is read-only.');
        }
        if ($propertyName === 'HasDataLossRisk') {
            throw new \RuntimeException('HasDataLossRisk is read-only.');
        }
    }

    public function Add(LoadIssue $issue)
    {
        $this->_issues[] = $issue;
    }
}