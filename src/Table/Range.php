<?php

namespace Slpi1\Excel\Table;

class Range
{
    public $start;
    public $end;

    public $active;

    public $maxX = 0;

    public function __construct(Cell $start, Cell $end)
    {
        $this->start = $start;
        $this->end   = $this->active   = $end;
    }

    public function haveCell(Cell $cell)
    {
        $cellX = $cell->getX();
        $cellY = $cell->getY();

        $startX = $this->start->getX();
        $startY = $this->start->getY();

        $endX = $this->end->getX();
        $endY = $this->end->getY();
        if ($cellX >= $startX && $cellX <= $endX && $cellY >= $startY && $cellY <= $endY) {
            return true;
        }
        return false;
    }

    public function isCell()
    {
        if ($this->start->getX() == $this->end->getX() && $this->start->getY() == $this->end->getY()) {
            return true;
        }
        return false;
    }

    public function horizonTouchAllow(Cell $cell)
    {
        if ($this->maxX != 0 && $cell->getX() > $this->maxX) {
            return false;
        }
        return true;
    }

    public function markMaxHorizonTouch()
    {
        $this->maxX = $this->maxX == 0 ? $this->active->getX() : min($this->maxX, $this->active->getX());
    }

    public function getValue()
    {
        return [
            [$this->start->getX(), $this->start->getY()],
            [$this->end->getX(), $this->end->getY()],
        ];
    }

    public function nextRowFirst()
    {
        $x = $this->start->getX();
        $y = $this->end->getY() + 1;
        return new Cell($x, $y);
    }

    public function nextColumnFirst()
    {
        $x = $this->end->getX() + 1;
        $y = $this->start->getY();
        return new Cell($x, $y);
    }

    public function nextMergeStartCells()
    {
        $nextRowFirst    = $this->nextRowFirst();
        $nextColumnFirst = $this->nextColumnFirst();
        return [
            $nextColumnFirst,
            $nextRowFirst,
        ];
    }

}
