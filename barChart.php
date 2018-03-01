<?php

//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels = [];

foreach ($_chart->seriesLabels as $s) {
    $dataSeriesLabels[] = new PHPExcel_Chart_DataSeriesValues(
            $_chart->seriesLabelsDataType, $sheet->name . '!' . $s, NULL, $_chart->seriesLabelsRowNumber
    );
}
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues = [];
foreach ($_chart->xAxisTickValues as $s) {
    $xAxisTickValues[] = new PHPExcel_Chart_DataSeriesValues(
            $_chart->xAxisTickValuesDataType, $sheet->name . '!' . $s, NULL, $_chart->xAxisTickValuesRowNumber
    );
}
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
/* $dataSeriesValues = array(
  new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', NULL, 4),
  new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', NULL, 4),
  new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$2:$D$5', NULL, 4),
  ); */
$dataSeriesValues = [];
foreach ($_chart->seriesValues as $s) {
    $dataSeriesValues[] = new PHPExcel_Chart_DataSeriesValues($_chart->seriesValuesDataType, $sheet->name . '!' . $s, NULL, $_chart->seriesValuesRowNumber);
}

//	Build the dataseries
$series = new PHPExcel_Chart_DataSeries(
        PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
        PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED, // plotGrouping
        range(0, count($dataSeriesValues) - 1), // plotOrder
        $dataSeriesLabels, // plotLabel
        $xAxisTickValues, // plotCategory
        $dataSeriesValues        // plotValues
);
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
$series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_VERTICAL);

//	Set the series in the plot area
$plotArea = new PHPExcel_Chart_PlotArea(NULL, array($series));
//	Set the chart legend
$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);

$title = new PHPExcel_Chart_Title($_chart->title);
$yAxisLabel = new PHPExcel_Chart_Title($_chart->valueTitle);


//	Create the chart
$chart = new PHPExcel_Chart(
        'chart1', // name
        $title, // title
        $legend, // legend
        $plotArea, // plotArea
        true, // plotVisibleOnly
        0, // displayBlanksAs
        NULL, // xAxisLabel
        $yAxisLabel  // yAxisLabel
);

//	Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition($_chart->topLeftPosition);
$chart->setBottomRightPosition($_chart->bottomRightPosition);

//	Add the chart to the worksheet
$activeSheet->addChart($chart);
