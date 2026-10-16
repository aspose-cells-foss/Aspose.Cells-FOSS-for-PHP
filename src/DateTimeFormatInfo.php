<?php

namespace Aspose\Cells_FOSS;

/**
 * Minimal DateTimeFormatInfo shim carrying date and time names and patterns.
 * The constructor populates the invariant culture; {@see DateTimeFormatInfo::SetCulture()}
 * replaces the values with localized data for the supported cultures.
 */
class DateTimeFormatInfo
{
    public $MonthNames;
    public $AbbreviatedMonthNames;
    public $DayNames;
    public $AbbreviatedDayNames;
    public $AMDesignator;
    public $PMDesignator;
    public $LongDatePattern;
    public $LongTimePattern;

    public function __construct()
    {
        $this->MonthNames = array(
            '',
            'January',
            'February',
            'March',
            'April',
            'May',
            'June',
            'July',
            'August',
            'September',
            'October',
            'November',
            'December',
            '',
        );
        $this->AbbreviatedMonthNames = array(
            '',
            'Jan',
            'Feb',
            'Mar',
            'Apr',
            'May',
            'Jun',
            'Jul',
            'Aug',
            'Sep',
            'Oct',
            'Nov',
            'Dec',
            '',
        );
        $this->DayNames = array(
            'Sunday',
            'Monday',
            'Tuesday',
            'Wednesday',
            'Thursday',
            'Friday',
            'Saturday',
        );
        $this->AbbreviatedDayNames = array(
            'Sun',
            'Mon',
            'Tue',
            'Wed',
            'Thu',
            'Fri',
            'Sat',
        );
        $this->AMDesignator = 'AM';
        $this->PMDesignator = 'PM';
        $this->LongDatePattern = 'dddd, dd MMMM yyyy';
        $this->LongTimePattern = 'HH:mm:ss';
    }

    public function SetCulture($name)
    {
        if (strcasecmp($name, 'en-US') === 0) {
            $this->AMDesignator = 'AM';
            $this->PMDesignator = 'PM';
            $this->LongDatePattern = 'dddd, MMMM d, yyyy';
            $this->LongTimePattern = 'h:mm:ss tt';
            return;
        }

        if (strcasecmp($name, 'fr-FR') === 0) {
            $this->MonthNames = array(
                '',
                'janvier',
                'février',
                'mars',
                'avril',
                'mai',
                'juin',
                'juillet',
                'août',
                'septembre',
                'octobre',
                'novembre',
                'décembre',
                '',
            );
            $this->AbbreviatedMonthNames = array(
                '',
                'janv.',
                'févr.',
                'mars',
                'avr.',
                'mai',
                'juin',
                'juil.',
                'août',
                'sept.',
                'oct.',
                'nov.',
                'déc.',
                '',
            );
            $this->DayNames = array(
                'dimanche',
                'lundi',
                'mardi',
                'mercredi',
                'jeudi',
                'vendredi',
                'samedi',
            );
            $this->AbbreviatedDayNames = array(
                'dim.',
                'lun.',
                'mar.',
                'mer.',
                'jeu.',
                'ven.',
                'sam.',
            );
            $this->AMDesignator = 'AM';
            $this->PMDesignator = 'PM';
            $this->LongDatePattern = 'dddd d MMMM yyyy';
            $this->LongTimePattern = 'HH:mm:ss';
            return;
        }

        if (strcasecmp($name, 'de-DE') === 0) {
            $this->MonthNames = array(
                '',
                'Januar',
                'Februar',
                'März',
                'April',
                'Mai',
                'Juni',
                'Juli',
                'August',
                'September',
                'Oktober',
                'November',
                'Dezember',
                '',
            );
            $this->AbbreviatedMonthNames = array(
                '',
                'Jan',
                'Feb',
                'Mär',
                'Apr',
                'Mai',
                'Jun',
                'Jul',
                'Aug',
                'Sep',
                'Okt',
                'Nov',
                'Dez',
                '',
            );
            $this->DayNames = array(
                'Sonntag',
                'Montag',
                'Dienstag',
                'Mittwoch',
                'Donnerstag',
                'Freitag',
                'Samstag',
            );
            $this->AbbreviatedDayNames = array(
                'So',
                'Mo',
                'Di',
                'Mi',
                'Do',
                'Fr',
                'Sa',
            );
            $this->AMDesignator = 'AM';
            $this->PMDesignator = 'PM';
            $this->LongDatePattern = 'dddd, d. MMMM yyyy';
            $this->LongTimePattern = 'HH:mm:ss';
            return;
        }

        if (strcasecmp($name, 'ja-JP') === 0) {
            $this->MonthNames = array(
                '',
                '1月',
                '2月',
                '3月',
                '4月',
                '5月',
                '6月',
                '7月',
                '8月',
                '9月',
                '10月',
                '11月',
                '12月',
                '',
            );
            $this->AbbreviatedMonthNames = $this->MonthNames;
            $this->DayNames = array(
                '日曜日',
                '月曜日',
                '火曜日',
                '水曜日',
                '木曜日',
                '金曜日',
                '土曜日',
            );
            $this->AbbreviatedDayNames = array(
                '日',
                '月',
                '火',
                '水',
                '木',
                '金',
                '土',
            );
            $this->AMDesignator = '午前';
            $this->PMDesignator = '午後';
            $this->LongDatePattern = 'yyyy\'年\'M\'月\'d\'日\'';
            $this->LongTimePattern = 'H:mm:ss';
            return;
        }
    }

    public function GetMonthName($month)
    {
        if ($month < 1 || $month > 13) {
            return '';
        }

        return $this->MonthNames[$month];
    }

    public function GetAbbreviatedMonthName($month)
    {
        if ($month < 1 || $month > 13) {
            return '';
        }

        return $this->AbbreviatedMonthNames[$month];
    }

    public function GetDayName($dayOfWeek)
    {
        if ($dayOfWeek < 0 || $dayOfWeek > 6) {
            return '';
        }

        return $this->DayNames[$dayOfWeek];
    }

    public function GetAbbreviatedDayName($dayOfWeek)
    {
        if ($dayOfWeek < 0 || $dayOfWeek > 6) {
            return '';
        }

        return $this->AbbreviatedDayNames[$dayOfWeek];
    }
}
