<?php

namespace Slpi1\Excel;

use Maatwebsite\Excel\Events\AfterSheet;
use Slpi1\Excel\Merge\Discover;
use Slpi1\Excel\Merge\StopRule;
use Slpi1\Excel\Repositories\WorksheetRepository;

class ExcelMergenable
{

    public static function afterSheet(AfterSheet $event)
    {
        $discover = new Discover(
            new WorksheetRepository($event->sheet),
            new StopRule($event->getConcernable())
        );

        $ranges = $discover->start();
        foreach ($ranges as $range) {
            $rangeName = $discover->getRangeName($range);
            $event->sheet->mergeCells($rangeName);
        }
    }
}
