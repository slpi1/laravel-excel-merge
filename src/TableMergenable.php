<?php

namespace Slpi1\Excel;

use Slpi1\Excel\Merge\Discover;
use Slpi1\Excel\Merge\StopRule;
use Slpi1\Excel\Repositories\ArrayRepository;
use Slpi1\Excel\Table\Cell;

class TableMergenable
{

    public function tbody($rule = null)
    {
        $discover = new Discover(
            new ArrayRepository($this->data),
            new StopRule(is_null($rule) ? $this : $rule)
        );

        $ranges = $discover->start();

        $body = [];
        foreach ($this->data as $rowIndex => $row) {
            $newRow = [
                'key' => $rowIndex,
            ];
            $i = 0;
            foreach ($row as $key => $value) {
                list($rowSpan, $colSpan) = $this->sumColumnAndRowSpan($i, $rowIndex, $ranges);
                $newRow[$key]            = [
                    'value'   => $value,
                    'colSpan' => $colSpan,
                    'rowSpan' => $rowSpan,
                ];
                $i++;
            }
            $body[] = $newRow;
        }
        return $body;
    }

    public function sumColumnAndRowSpan($x, $y, $ranges)
    {
        $rowSpan = $colSpan = 1;
        foreach ($ranges as $range) {
            $cell = new Cell($x, $y);

            if ($range->haveCell($cell)) {
                if ($range->start->equal($cell)) {
                    $rowSpan = $range->end->getY() - $range->start->getY() + 1;
                    $colSpan = $range->end->getX() - $range->start->getX() + 1;
                } else {
                    $rowSpan = $colSpan = 0;
                }
                break;
            }
        }
        return [$rowSpan, $colSpan];
    }
}
