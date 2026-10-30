<?php
namespace Aspose\Cells_FOSS\Core;

final class DateSerialConverter
{
    private static $Windows1900Epoch;
    private static $Mac1904Epoch;
    const TicksPerDay = 864000000000;

    private static function initEpochs()
    {
        if (self::$Windows1900Epoch === null) {
            self::$Windows1900Epoch = new \DateTime("1899-12-31 00:00:00");
        }
        if (self::$Mac1904Epoch === null) {
            self::$Mac1904Epoch = new \DateTime("1904-01-01 00:00:00");
        }
    }

    public static function ToSerial($value, $dateSystem)
    {
        self::initEpochs();
        $baseDate = ($dateSystem === DateSystem::Mac1904) ? self::$Mac1904Epoch : self::$Windows1900Epoch;
        $serial = (self::getTicks($value) - self::getTicks($baseDate)) / self::TicksPerDay;
        if ($dateSystem === DateSystem::Windows1900 && $serial >= 60.0) {
            $serial += 1.0;
        }
        return $serial;
    }

    public static function FromSerial($serial, $dateSystem)
    {
        self::initEpochs();
        $s = $serial;
        if ($dateSystem === DateSystem::Windows1900 && $s >= 60.0) {
            $s -= 1.0;
        }
        $baseDate = ($dateSystem === DateSystem::Mac1904) ? self::$Mac1904Epoch : self::$Windows1900Epoch;
        $ticks = (int)round($s * self::TicksPerDay);
        $result = clone $baseDate;
        $days = intdiv($ticks, self::TicksPerDay);
        $rem = $ticks % self::TicksPerDay;
        if ($rem !== 0) {
            $result->modify(($days >= 0 ? "+" : "") . $days . " day");
            $seconds = intdiv($rem, 10000000);
            if ($seconds !== 0) {
                $result->modify(($seconds >= 0 ? "+" : "") . $seconds . " second");
            }
        } else if ($days !== 0) {
            $result->modify(($days >= 0 ? "+" : "") . $days . " day");
        }
        return $result;
    }

    private static function getTicks($dateTime)
    {
        $timestamp = $dateTime->getTimestamp();
        $microseconds = (int) $dateTime->format('u');
        return ($timestamp + 62135596800) * 10000000 + $microseconds * 10;
    }
}
