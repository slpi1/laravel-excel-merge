<?php

namespace Slpi1\Excel\Concerns;

use Slpi1\Excel\Table\Cell;

interface RepositoryInterface
{
    public function getValue(Cell $cell);
    public function getWidth();
    public function getHeight();
}
