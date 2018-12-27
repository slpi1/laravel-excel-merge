<?php

namespace Slpi1\Excel\Repositories;

use Slpi1\Excel\Concerns\RepositoryInterface;
use Slpi1\Excel\Table\Cell;

class ArrayRepository implements RepositoryInterface
{

    private $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function getValue(Cell $cell)
    {
        return $this->data[$cell->getY()][$cell->getX()];
    }

    public function getWidth()
    {
        return count($this->data[0]);
    }

    public function getHeight()
    {
        return count($this->data);
    }
}
