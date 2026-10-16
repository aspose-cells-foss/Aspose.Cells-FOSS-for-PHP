<?php

namespace Aspose\Cells_FOSS;

/**
 * Minimal NumberFormatInfo shim carrying the invariant culture's numeric
 * separators, group sizes and sign symbols. Localized data is applied by
 * {@see CultureInfo} through {@see NumberFormatInfo::SetCulture()}.
 */
class NumberFormatInfo
{
    public $NumberDecimalSeparator;
    public $NumberGroupSeparator;
    public $NumberGroupSizes;
    public $PercentSymbol;
    public $NegativeSign;
    public $PositiveSign;

    public function __construct()
    {
        $this->NumberDecimalSeparator = '.';
        $this->NumberGroupSeparator = ',';
        $this->NumberGroupSizes = array(3);
        $this->PercentSymbol = '%';
        $this->NegativeSign = '-';
        $this->PositiveSign = '';
    }

    public function SetCulture($name)
    {
        if (strcasecmp($name, 'en-US') === 0) {
            $this->NumberDecimalSeparator = '.';
            $this->NumberGroupSeparator = ',';
            $this->NumberGroupSizes = array(3);
            $this->PercentSymbol = '%';
            $this->NegativeSign = '-';
            $this->PositiveSign = '';
            return;
        }

        if (strcasecmp($name, 'fr-FR') === 0) {
            $this->NumberDecimalSeparator = ',';
            $this->NumberGroupSeparator = "\u{202F}";
            $this->NumberGroupSizes = array(3);
            $this->PercentSymbol = '%';
            $this->NegativeSign = '-';
            $this->PositiveSign = '';
            return;
        }

        if (strcasecmp($name, 'de-DE') === 0) {
            $this->NumberDecimalSeparator = ',';
            $this->NumberGroupSeparator = '.';
            $this->NumberGroupSizes = array(3);
            $this->PercentSymbol = '%';
            $this->NegativeSign = '-';
            $this->PositiveSign = '';
            return;
        }

        if (strcasecmp($name, 'ja-JP') === 0) {
            $this->NumberDecimalSeparator = '.';
            $this->NumberGroupSeparator = ',';
            $this->NumberGroupSizes = array(3);
            $this->PercentSymbol = '%';
            $this->NegativeSign = '-';
            $this->PositiveSign = '';
            return;
        }
    }
}
