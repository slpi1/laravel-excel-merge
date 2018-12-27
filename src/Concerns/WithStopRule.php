<?php

namespace Slpi1\Excel\Concerns;

interface WithStopRule
{
    /**
     * @return  boolean     true | false
     *          array       ['B', 'D', 'H'] | [0,1,2]
     *          Closure     function($x, $y)
     */
    public function stopColumns();

    /**
     * @return  boolean     true | false
     *          array       [0,1,2]
     *          Closure     function($y, $x)
     */
    public function stopRows();
}
