<?php

namespace Slpi1\Excel\Table;

class Cell
{

    protected $x;
    protected $y;

    public function __construct($x = 0, $y = 0)
    {
        $this->x = $x;
        $this->y = $y;
    }

    public function getX()
    {
        return $this->x;
    }
    public function getY()
    {
        return $this->y;
    }

    public function nextHorizonCell()
    {
        return new Cell($this->x + 1, $this->y);
    }

    public function nextVerticalCell()
    {
        return new Cell($this->x, $this->y + 1);
    }

    public function equal(Cell $cell)
    {
        return $this->x == $cell->getX() && $this->y == $cell->getY();
    }
}
