<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel.php';

// 

$objPHPExcel = new PHPExcel();

$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("PHPExcel Test Document")
							 ->setSubject("PHPExcel Test Document")
							 ->setDescription("Test document for PHPExcel, generated using PHP classes.")
							 ->setKeywords("office PHPExcel php")
							 ->setCategory("Test result file");

 function exist($val, $default){
	if(!isset($val)){
		return $default;
	}
	return $val;
}
							
$file = $argv[1]; //json file 
							
$json =  json_decode(file_get_contents($file));
					

$filePath  =  $json->path ?? 'file.xlsx';
							
$sheets  = $json->sheets?? [];

unlink($file);


$sheetCounter = 0;
foreach($sheets as $sheet){

	
$datas  = $sheet->cells ?? [];
$colWidth = $sheet->columns?? [];
$drawings = $sheet->drawings?? [];
$merges = $sheet->merges ?? [];



//if not default sheet create a new one
if($sheetCounter !=0 ){
	$objPHPExcel->createSheet();
}
$activeSheet = $objPHPExcel->setActiveSheetIndex($sheetCounter);

foreach($colWidth as $key=>$col){
	if(isset($col->width)){
		$activeSheet->getColumnDimension($key)->setWidth($col->width);
	}
}

foreach($merges as $key){
	$activeSheet->mergeCells($key);
}
foreach($datas as $key=>$cell){

		$activeSheet ->setCellValue($key, $cell->val);

		if(isset($cell->underline)){
			$activeSheet->getStyle($key)->getFont()->setUnderline($cell->underline);
		}


		if(isset($cell->styles)){
			
			//cell's type

			if(isset($cell->styles->format)){
				$type = $cell->styles->format??'string';
				if($type=='date'){
					$dateTime = new DateTime($cell->val); 
					$activeSheet->setCellValue($key, $dateTime);
					$activeSheet->getStyle($key)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY);
				}else{
					$activeSheet->getStyle($key)->getNumberFormat()->setFormatCode($cell->styles->format);
					
				}
			}

			//cell's font size
			if(isset($cell->styles->fontSize)){
				$activeSheet->getStyle($key) ->getFont()->setSize($cell->styles->fontSize);
			}
			//cell's bgcolor (a RGB color)
			if(isset($cell->styles->fill)){
				$activeSheet->getStyle($key) ->getFill()
				->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				->getStartColor()
				->setRGB($cell->styles->fill);
			}
			//cell border
			if(isset($cell->styles->border)){
				$style = $cell->styles->border->style;
				$excel_boder_style = '';
				switch ($style) {
					case 'thin':
					$excel_boder_style = PHPExcel_Style_Border::BORDER_THIN;
						break;
					
					default:
					$excel_boder_style = PHPExcel_Style_Border::BORDER_THICK;
						break;
				}
				$color = $cell->styles->border->color;
				$position =  @$cell->styles->border->position;
				$activeSheet->getStyle($key)->applyFromArray(
					$border_style= array('borders' => array(exist($position, 'allborders') => array('style' => 
					$excel_boder_style,'color' => array('argb' => $color),)))
				);
			}

			//font
			if(isset($cell->styles->font)) {
				$styleArray = array('font'  => json_decode(json_encode($cell->styles->font), True));
				$activeSheet->getStyle($key)->applyFromArray($styleArray);
			}
			//alignment and rotation
			if(isset($cell->styles->alignment)){
				$_key = $cell->styles->alignment->key??'horizontal';
				$value = $cell->styles->alignment->value??'left';
				
				$rotation = $cell->styles->alignment->rotation ?? 0;
				$styleArray = array(
					'alignment'=>array(
						$_key => $value
					)
				);
				$activeSheet->getStyle($key)->applyFromArray($styleArray);
				$activeSheet->getStyle($key)->getAlignment()->setTextRotation( exist($rotation, 0));
			}
		
    }
}

//including pictures
foreach($drawings as $dr){
   
	$objDrawing = new PHPExcel_Worksheet_Drawing();
	$objDrawing->setWorksheet($activeSheet);
	$objDrawing->setName($dr->name);
	$objDrawing->setDescription(exist(@$dr->description, $dr->name));
	$objDrawing->setPath($dr->path);
	$objDrawing->setCoordinates($dr->cordonnates);
	$objDrawing->setOffsetX($dr->offsetX);
	$objDrawing->setHeight($dr->height);
	$objDrawing->setRotation(exist(@$dr->rotation, 0));
	$objDrawing->getShadow()->setVisible(exist(@$dr->visible, true));
	$objDrawing->getShadow()->setDirection(exist(@$dr->direction, 45));
}


$objPHPExcel->getActiveSheet()->setTitle($sheet->name);




if(count($sheet->charts) > 0){
	foreach($sheet->charts as $_chart){
		if($_chart->type=='barChart'){
			require_once(__DIR__."\\barChart.php");
		}
	}
}
$sheetCounter++;
}

$objPHPExcel->setActiveSheetIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$objWriter->save($filePath);


// echo file_get_contents($filePath);