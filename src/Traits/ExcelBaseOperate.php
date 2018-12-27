<?php

namespace Slpi1\Excel\Traits;

use Slpi1\Excel\Table\Cell;
use Slpi1\Excel\Table\Range;

trait ExcelBaseOperate
{
    protected $letters = [
        'A' => 1,
        'B' => 2,
        'C' => 3,
        'D' => 4,
        'E' => 5,
        'F' => 6,
        'G' => 7,
        'H' => 8,
        'I' => 9,
        'J' => 10,
        'K' => 11,
        'L' => 12,
        'M' => 13,
        'N' => 14,
        'O' => 15,
        'P' => 16,
        'Q' => 17,
        'R' => 18,
        'S' => 19,
        'T' => 20,
        'U' => 21,
        'V' => 22,
        'W' => 23,
        'X' => 24,
        'Y' => 25,
        'Z' => 26,
    ];

    public function isColumnName($name)
    {
        return preg_match('/^[A-Za-z]+$/', $name);
    }

    public function indexToColumnName($index)
    {
        $name = '';
        $list = $this->getNumList($index);
        foreach ($list as $num) {
            $name .= array_search($num, $this->letters);
        }
        return $name;
    }

    private function getNumList($index, $stick = [])
    {

        $divisor = (int) floor($index / 26);
        $remain  = $index % 26 + 1;
        array_unshift($stick, $remain);
        if ($divisor > 26) {
            return $this->getNumList($divisor);
        } else if ($divisor > 0) {
            array_unshift($stick, $divisor);
            return $stick;
        } else {
            return $stick;
        }
    }

    public function columnNameToIndex($columnName)
    {
        $columnName = strtoupper($columnName);

        $index = 0;
        $len   = strlen($columnName) - 1;
        for ($i = 0; $len >= 0; $len--, $i++) {
            $letter = $columnName[$len];

            $index += $this->letters[$letter] * pow(26, $i);
        }
        return $index - 1;
    }

    public function getRangeId(Range $range)
    {
        return $this->getCellId($range->start);
    }

    public function getRangeName(Range $range)
    {
        return $this->getCellId($range->start) . ':' . $this->getCellId($range->end);
    }

    public function getCellId(Cell $cell)
    {
        return $this->getCellHorizonIndex($cell);
    }

    public function getCellHorizonIndex(Cell $cell)
    {
        return $this->indexToColumnName($cell->getX()) . ($cell->getY() + 1);
    }

    public function getCellVerticalIndex(Cell $cell)
    {
        return $this->indexToColumnName($cell->getY()) . ($cell->getX() + 1);
    }
}
