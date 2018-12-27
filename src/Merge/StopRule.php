<?php

namespace Slpi1\Excel\Merge;

use Slpi1\Excel\Concerns\WithStopRule;
use Slpi1\Excel\Table\Cell;
use Slpi1\Excel\Traits\ExcelBaseOperate;

class StopRule
{
    use ExcelBaseOperate;

    private $exporter;
    protected $stopColumn = [];
    protected $stopRow    = [];

    public $extendStopRows = false;

    public $extendStopColumns = false;

    public function __construct($exporter)
    {
        $this->exporter = $exporter;

        if (isset($exporter->extendStopRows) && $exporter->extendStopRows) {
            $this->extendStopRows = true;
        }

        if (isset($exporter->extendStopColumns) && $exporter->extendStopColumns) {
            $this->extendStopColumns = true;
        }
    }

    public function atStopColumn(Cell $cell)
    {
        if ($this->exporter instanceof WithStopRule) {
            $rule       = $this->exporter->stopColumns();
            $shouldStop = $this->parserStopColumnRule($rule, $cell);

            if ($shouldStop) {
                return true;
            }
        }

        if (in_array($cell->getX(), $this->stopColumn)) {
            return true;
        }
        return false;
    }

    public function atStopRow(Cell $cell)
    {
        if ($this->exporter instanceof WithStopRule) {
            $rule       = $this->exporter->stopRows();
            $shouldStop = $this->parserStopRowRule($rule, $cell);

            if ($shouldStop) {
                return true;
            }
        }

        if (in_array($cell->getY(), $this->stopRow)) {
            return true;
        }
        return false;
    }

    public function addStopColumn($index)
    {
        $this->stopColumn[$index] = $index;
    }

    public function addStopRow($index)
    {
        $this->stopRow[$index] = $index;
    }

    public function parserStopColumnRule($rule, Cell $cell)
    {
        if (is_bool($rule)) {
            return $rule;
        }

        if (is_array($rule)) {
            foreach ($rule as $item) {
                if (is_numeric($item) && $cell->getX() == $item) {
                    return true;
                }

                if ($this->isColumnName($item) && $this->columnNameToIndex($item) == $cell->getX()) {
                    return true;
                }
            }
        }

        if (is_callable($rule)) {
            return call_user_func_array($rule, [$cell->getX(), $cell->getY()]);
        }

        return false;
    }

    public function parserStopRowRule($rule, Cell $cell)
    {
        if (is_bool($rule)) {
            return $rule;
        }

        if (is_array($rule)) {
            foreach ($rule as $item) {
                if (is_numeric($item) && $cell->getY() == $item) {
                    return true;
                }
            }
        }

        if (is_callable($rule)) {
            return call_user_func_array($rule, [$cell->getY(), $cell->getX()]);
        }

        return false;
    }
}
