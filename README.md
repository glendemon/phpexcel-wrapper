# phpexcel-wrapper
Class wrapper for PHPExcel

# Examples

```php
        $chart = PhpExcelWrapper::getChart(
                $data,
                PHPExcel_Chart_DataSeries::TYPE_LINECHART,
                PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
                null,
                '[h]:mm:ss'
            );

        $chart1 = PhpExcelWrapper::getChart(
                $data,
                PHPExcel_Chart_DataSeries::TYPE_BARCHART,
                PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,
                $columns,
                '[h]:mm:ss'
            );

        $chart2 = PhpExcelWrapper::getChart(
                $data,
                PHPExcel_Chart_DataSeries::TYPE_BARCHART,
                PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,
                $columns,
                null,
                20
            );

        $xls = new PhpExcelWrapper($dirName);
        if (is_array($function)) {
			$xls->setData($function, $titles, $chart);
		} else {
			$xls->splitToSheets($function, $titles, $chart);
		}
		if (!$download) {
			$xls->save();
		} else {
			$xls->downloadFromStream();
		}
```
