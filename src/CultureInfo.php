<?php

namespace Aspose\Cells_FOSS;

/**
 * Minimal CultureInfo shim.
 *
 * The .NET {@see \System.Globalization.CultureInfo} type carries a full
 * locale database. This port models the parts used by the library: the
 * invariant culture plus a small set of supported cultures that carry their
 * own date/time and number formatting data. Locale lookups via
 * {@see CultureInfo::GetCultureInfo()} accept a culture name or an LCID and
 * fall back to invariant formatting data for unknown cultures.
 */
class CultureInfo
{
    public static $InvariantCulture = null;

    public $Name;
    public $LCID;
    public $DateTimeFormat;
    public $NumberFormat;

    public function __construct($name = '', $lcid = 127)
    {
        $this->Name = $name;
        $this->LCID = $lcid;
        $this->DateTimeFormat = new DateTimeFormatInfo();
        $this->NumberFormat = new NumberFormatInfo();
        if ($name === null || trim($name) === '') {
            return;
        }

        $this->DateTimeFormat->SetCulture($name);
        $this->NumberFormat->SetCulture($name);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'InvariantCulture') {
            return self::GetInvariantCulture();
        }

        return null;
    }

    public static function GetInvariantCulture()
    {
        static $invariant = null;
        if ($invariant === null) {
            $invariant = new CultureInfo('', 127);
        }

        return $invariant;
    }

    public static function GetCultureInfo($lcid)
    {
        if (is_int($lcid) && $lcid === 127) {
            return self::GetInvariantCulture();
        }

        $name = self::NormalizeCultureName($lcid);
        if ($name === null || $name === '') {
            return self::GetInvariantCulture();
        }

        return new CultureInfo($name, self::ResolveLcid($name));
    }

    private static function NormalizeCultureName($lcid)
    {
        if (is_int($lcid)) {
            if ($lcid === 127) {
                return '';
            }

            if ($lcid === 1033) {
                return 'en-US';
            }

            if ($lcid === 1036) {
                return 'fr-FR';
            }

            if ($lcid === 1031) {
                return 'de-DE';
            }

            if ($lcid === 1041) {
                return 'ja-JP';
            }

            return null;
        }

        $trimmed = trim((string) $lcid);
        if ($trimmed === '') {
            return '';
        }

        if (preg_match('/^[0-9]+$/', $trimmed) === 1) {
            return self::NormalizeCultureName((int) $trimmed);
        }

        if (strcasecmp($trimmed, 'en-US') === 0 || strcasecmp($trimmed, 'en') === 0) {
            return 'en-US';
        }

        if (strcasecmp($trimmed, 'fr-FR') === 0 || strcasecmp($trimmed, 'fr') === 0) {
            return 'fr-FR';
        }

        if (strcasecmp($trimmed, 'de-DE') === 0 || strcasecmp($trimmed, 'de') === 0) {
            return 'de-DE';
        }

        if (strcasecmp($trimmed, 'ja-JP') === 0 || strcasecmp($trimmed, 'ja') === 0) {
            return 'ja-JP';
        }

        return null;
    }

    private static function ResolveLcid($name)
    {
        if (strcasecmp($name, 'en-US') === 0) {
            return 1033;
        }

        if (strcasecmp($name, 'fr-FR') === 0) {
            return 1036;
        }

        if (strcasecmp($name, 'de-DE') === 0) {
            return 1031;
        }

        if (strcasecmp($name, 'ja-JP') === 0) {
            return 1041;
        }

        return 127;
    }

    public function Clone()
    {
        $clone = new CultureInfo($this->Name, $this->LCID);
        $clone->DateTimeFormat = clone $this->DateTimeFormat;
        $clone->NumberFormat = clone $this->NumberFormat;
        return $clone;
    }
}

CultureInfo::$InvariantCulture = new CultureInfo('', 127);
