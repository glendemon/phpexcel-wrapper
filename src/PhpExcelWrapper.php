<?php

/**
 * PhpExcelWrapper.
 */
class PhpExcelWrapper
{
    /**
     * Used to split to worksheets.
     * @var int
     */
    protected $limitLength = 1000;

	/**
	 * @var string
	 */
    protected $filename;

	/**
	 * @var \PHPExcel
	 */
    protected $objPHPExcel;

    /**
     * @param string $dirName Target directory for xls files
     */
    public function __construct($dirName)
    {
        $cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '512MB');
        PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
        $this->filename = realpath($dirName . '/') . date("dmY-His") . "_" . uniqid() . ".xlsx";
        $this->objPHPExcel = new PHPExcel();
    }

    /**
     * Save to file.
     *
     * @return string filename
     */
    public function save()
    {
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
        $objWriter->setIncludeCharts(TRUE);
        $objWriter->save($this->filename);
        return $this->filename;
    }

    /**
     * Set data to active sheet.
     * 
     * @param array $data
     * @param array $title
     * @param PHPExcel_Chart $charts
     */
    public function setData($data, $title = array(), $charts = null)
    {
        if (!empty($title)) {
			array_unshift($data, $title);
		}

        PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_RCUPValueBinder() );
        $this->objPHPExcel->getActiveSheet()->fromArray($data);

        if (!empty($charts))
        {
            if (!is_array($charts)) {
				$charts = array($charts);
			}

			foreach ($charts as $chart)
            {
                $this->objPHPExcel->getActiveSheet()->addChart($chart);
            }
        }
    }

    /**
     * Set data to active sheet and create new one.
     * 
     * @param array $data
     * @param array $title
     * @param PHPExcel_Chart $charts
     */
    public function newSheet($data, $title = array(), $charts = null)
    {
        $this->setData($data, $title);
        $this->objPHPExcel->createSheet();
        $this->objPHPExcel->setActiveSheetIndex($this->objPHPExcel->getSheetCount() - 1);
    }

    /**
     * Retrieve data and split to sheets.
     * Example:
     * <pre><code>
     * <?php
        $function = function($offset, $limit)
        {
            $result = ...
            return array(
                $result['data'],
                $result['total'],
            );
        };
        $xls = new XLSReport();
        $xls->splitToSheets($function, $titles);
     * ?>
     * </code></pre>
     * @param function $function
     * @param array $titles
     * @param PHPExcel_Chart $charts
     */
    public function splitToSheets(&$function, &$titles = array(), $charts = null)
    {
        $offset = 0;
        $limit = $this->limitLength;
        do {
            list($data, $total) = $function($offset, $limit);
            if (!empty($data)) {
				$this->newSheet($data, $titles, $charts);
			}
			$offset += $limit;
        }
        while ($offset < $total);
        $this->objPHPExcel->removeSheetByIndex($this->objPHPExcel->getSheetCount() - 1);
    }

    /**
     * Download saved file.
     */
    public function download()
    {
        $this->SendHeaders();
        readfile($this->filename);
    }

    /**
     * Download file without save.
     */
    public function downloadFromStream()
    {
        $this->SendHeaders();
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
        $objWriter->setIncludeCharts(TRUE);
        $objWriter->save('php://output');
    }

    /**
     * Send headers for download.
     */
    protected function SendHeaders()
    {
        header("Content-Description: File Transfer\r\n");
        header("Pragma: public\r\n");
        header("Expires: 0\r\n");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0\r\n");
        header("Cache-Control: public\r\n");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment; filename=\"" . basename($this->filename) . "\"\r\n");
    }

    /**
     * Convert xls file to array
     * 
     * @param string $filepath
     * @return array
     */
    public static function XLSToArray($filepath)
    {
        $result = array();
        try {
            if (file_exists($filepath))
            {
                $file_extension = end(explode(".", $filepath));

                if ($file_extension == "xlsx") {
					$objReader = PHPExcel_IOFactory::createReader('Excel2007');
				} else {
					$objReader = PHPExcel_IOFactory::createReader('Excel5');
				}

				$objPHPExcel = $objReader->load($filepath);
                $objPHPExcel->setActiveSheetIndex(0);
                $result = $objPHPExcel->setActiveSheetIndex(0)->toArray();
            }
        }
        catch (Exception $e) {
            return array();
        }
        return $result;
    }

    /**
     *
     * @param array $data
     * @param $plotType PHPExcel_Chart_DataSeries::TYPE_*
     * @param $plotGrouping PHPExcel_Chart_DataSeries::GROUPING_*
     * @param array $columns
     * @param string $format format for values
     * @param int $positionOffset
     * @return \PHPExcel_Chart
     */
    public static function getChart(
        $data,
        $plotType,
        $plotGrouping,
        $columns = null,
        $format = null,
        $positionOffset = 0
    ) {
        $columnCount = count($data[0]);
        $rowCount = count($data);
        $keys = array_keys($data[0]);
        $labels = array();
        $categories = array();
        $values = array();
        for ($i = 1; $i < $columnCount; $i++)
        {
            if (!is_array($columns) || in_array($keys[$i], $columns))
            {
                $col = PHPExcel_Cell::stringFromColumnIndex($i);
                $labels[] = new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$'.$col.'$1', null, 1);
                $categories[] = new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$'.($rowCount + 1), null, $rowCount);
                $values[] = new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$'.$col.'$2:$'.$col.'$'.($rowCount + 1), $format, $rowCount);
            }
        }
        $series = new PHPExcel_Chart_DataSeries(
          $plotType,      // plotType
          $plotGrouping,   // plotGrouping
          range(0, count($values)-1),                     // plotOrder
          $labels,                                        // plotLabel
          $categories,                                    // plotCategory
          $values                                         // plotValues
        );
        $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
        $plotarea = new PHPExcel_Chart_PlotArea(null, array($series));
        $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, null, false);
        $chart = new PHPExcel_Chart(
          'chart'.uniqid(),                               // name
          null,                                           // title
          $legend,                                        // legend
          $plotarea,                                      // plotArea
          true,                                           // plotVisibleOnly
          0,                                              // displayBlanksAs
          null,                                           // xAxisLabel
          null                                            // yAxisLabel
        );
        $chart->setTopLeftPosition(PHPExcel_Cell::stringFromColumnIndex($columnCount+1).''.(2+$positionOffset));
        $chart->setBottomRightPosition(PHPExcel_Cell::stringFromColumnIndex($columnCount+19).''.(20+$positionOffset));

        return $chart;
    }

    /**
     * Сортируем ассоциативный массив так, чтобы порядок столбцов соответствовал порядку тайтлов сверху
     * 
     * @param array $data
     * @param array $titles
     * @return array
     */
    public static function makeTheSameOrder($data, $titles) {
        $count = count($data);
        $result = array();
        foreach($titles as $title) {
            for($i = 0; $i < $count; $i++) {
                $result[$i][$title] = $data[$i][$title];
            }
        }
        return $result;
    }

}
