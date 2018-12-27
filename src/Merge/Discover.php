<?php

namespace Slpi1\Excel\Merge;

use Slpi1\Excel\Concerns\RepositoryInterface;
use Slpi1\Excel\Table\Cell;
use Slpi1\Excel\Table\Range;
use Slpi1\Excel\Traits\ExcelBaseOperate;

class Discover
{
    use ExcelBaseOperate;

    const HORIZON_DIRECT  = 0;
    const VERTICAL_DIRECT = 1;

    private $repository;

    protected $stopRule;

    // 优先合并方向
    protected $mergeDirect;

    // 已合并区域
    protected $mergedRanges = [];

    // 横向合并队列
    private $horizonMergeQueue;

    // 纵向合并队列
    private $verticalMergeQueue;

    private $width;

    private $height;

    public function __construct(RepositoryInterface $repository, $stopRule)
    {
        $this->repository = $repository;
        $this->stopRule   = $stopRule;

        $this->height = $this->repository->getHeight();
        $this->width  = $this->repository->getWidth();

        if ($this->height > $this->width) {
            $this->mergeDirect = self::HORIZON_DIRECT;
        } else {
            $this->mergeDirect = self::VERTICAL_DIRECT;
        }
    }

    public function start()
    {
        $this->horizonMergeQueue[] = new Cell;
        $this->catchMergeRange();
        return $this->mergedRanges;
    }

    public function catchMergeRange()
    {
        $startCell = $this->getNextMergeStart();
        if ($startCell) {
            if (!$this->merged($startCell)) {
                $range = new Range($startCell, $startCell);

                $range = $this->touchRange($range);
                if (!$range->isCell()) {
                    $this->mergedRanges[$this->getRangeId($range)] = $range;
                }

                list($cellHorizon, $cellVertical) = $range->nextMergeStartCells();

                if ($this->isExistsCell($cellHorizon)) {
                    $this->horizonMergeQueue[$this->getCellHorizonIndex($cellHorizon)] = $cellHorizon;
                    $this->makeExtendStopRule($cellHorizon);
                }

                if ($this->isExistsCell($cellVertical)) {
                    $this->verticalMergeQueue[$this->getCellVerticalIndex($cellVertical)] = $cellVertical;
                    $this->makeExtendStopRule($cellVertical);
                }
            }

            $this->catchMergeRange();
        }
    }

    public function merged(Cell $cell)
    {
        if (isset($this->mergedRanges[$this->getCellId($cell)])) {
            return true;
        }

        foreach ($this->mergedRanges as $range) {
            if ($range->haveCell($cell)) {
                return true;
            }
        }
        return false;
    }

    public function touchRange(Range $range)
    {
        $range = $this->touchCellHorizon($range);
        $range = $this->touchCellVertical($range);
        return $range;
    }

    public function touchCellHorizon(Range $range)
    {
        $nextHorizonCell = $range->active->nextHorizonCell();
        if (!$this->isExistsCell($nextHorizonCell)) {
            $range->end = $range->active;
            return $range;
        }

        if ($range->horizonTouchAllow($nextHorizonCell) &&
            !$this->atStopColumn($nextHorizonCell) &&
            $this->getValue($range->active) === $this->getValue($nextHorizonCell)) {
            $range->active = $nextHorizonCell;
            return $this->touchCellHorizon($range);
        } else {
            $range->end = $range->active;

            $range->markMaxHorizonTouch();
        }

        return $range;
    }

    public function touchCellVertical(Range $range)
    {
        $nextRowFirst = $range->nextRowFirst();
        if (!$this->isExistsCell($nextRowFirst)) {
            return $range;
        }

        if (!$this->atStopRow($nextRowFirst) && $this->getValue($range->start) === $this->getValue($nextRowFirst)) {
            $range->active = $nextRowFirst;
            return $this->touchRange($range);
        }
        return $range;
    }

    public function getNextMergeStart()
    {
        if (empty($this->horizonMergeQueue) && empty($this->verticalMergeQueue)) {
            return false;
        }

        if (($this->mergeDirect == self::HORIZON_DIRECT && !empty($this->horizonMergeQueue)) || empty($this->verticalMergeQueue)) {
            return $this->getNextMergeStartOfQueue($this->horizonMergeQueue);
        } else {
            return $this->getNextMergeStartOfQueue($this->verticalMergeQueue);
        }
    }

    public function getNextMergeStartOfQueue(&$queue)
    {
        krsort($queue);
        return array_pop($queue);
    }

    public function getValue(Cell $cell)
    {
        return $this->repository->getValue($cell);
    }

    // 是否是停止列
    public function atStopColumn(Cell $cell)
    {
        return $this->stopRule->atStopColumn($cell);
    }

    // 是否是停止行
    public function atStopRow(Cell $cell)
    {
        return $this->stopRule->atStopRow($cell);
    }

    public function isExistsCell(Cell $cell)
    {
        return $cell->getX() < $this->width && $cell->getY() < $this->height;
    }

    public function makeExtendStopRule(Cell $cell)
    {
        if ($this->stopRule->extendStopRows) {
            $this->stopRule->addStopRow($cell->getY());
        }
        if ($this->stopRule->extendStopColumns) {
            $this->stopRule->addStopColumn($cell->getX());
        }
    }
}
