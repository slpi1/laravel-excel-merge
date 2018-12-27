<?php

namespace Slpi1\Excel\Repositories;

use Maatwebsite\Excel\Sheet;
use Slpi1\Excel\Concerns\RepositoryInterface;
use Slpi1\Excel\Table\Cell;
use Slpi1\Excel\Traits\ExcelBaseOperate;

class WorksheetRepository implements RepositoryInterface
{
    use ExcelBaseOperate;

    private $sheet;

    public function __construct(Sheet $sheet)
    {
        $this->sheet = $sheet;
    }

    public function getValue(Cell $cell)
    {
        return $this->sheet->getCell($this->getCellId($cell))->getValue();
    }

    public function getWidth()
    {
        return $this->columnNameToIndex($this->sheet->getHighestColumn()) + 1;
    }

    public function getHeight()
    {
        return (int) $this->sheet->getHighestRow();
    }
}
