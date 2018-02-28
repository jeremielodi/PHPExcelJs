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
foreach($datas as $key=>$data){
    
		$activeSheet ->setCellValue($key, $data->val);
		if(isset($data->styles)){
			
			//cell's type
			if(isset($data->styles->format)){
				$activeSheet->getStyle($key)->getNumberFormat()->setFormatCode($data->styles->format);
			}

			//cell's font size
			if(isset($data->styles->fontSize)){
				$activeSheet->getStyle($key) ->getFont()->setSize($data->styles->fontSize);
			}
			//cell's bgcolor (a RGB color)
			if(isset($data->styles->fill)){
				$activeSheet->getStyle($key) ->getFill()
				->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				->getStartColor()
				->setRGB($data->styles->fill);
			}
			//cell border
			if(isset($data->styles->border)){
				$style = $data->styles->border->style;
				$excel_boder_style = '';
				switch ($style) {
					case 'thin':
					$excel_boder_style = PHPExcel_Style_Border::BORDER_THIN;
						break;
					
					default:
					$excel_boder_style = PHPExcel_Style_Border::BORDER_THICK;
						break;
				}
				$color = $data->styles->border->color;
				$position =  @$data->styles->border->position;
				$activeSheet->getStyle($key)->applyFromArray(
					$border_style= array('borders' => array(exist($position, 'allborders') => array('style' => 
					$excel_boder_style,'color' => array('argb' => $color),)))
				);
			}

			//font
			if(isset($data->styles->font)) {
				$styleArray = array('font'  => json_decode(json_encode($data->styles->font), True));
				$activeSheet->getStyle($key)->applyFromArray($styleArray);
			}
			//alignment and rotation
			if(isset($data->styles->alignment)){
				$_key = $data->styles->alignment->key;
				$value = $data->styles->alignment->value;
				
				$rotation = $data->styles->alignment->rotation ?? 0;
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



//$objPHPExcel->addNamedRange( new PHPExcel_NamedRange('PersonFN', $objPHPExcel->getActiveSheet(), 'B1') );
//$objPHPExcel->addNamedRange( new PHPExcel_NamedRange('PersonLN', $objPHPExcel->getActiveSheet(), 'B2') );

$objPHPExcel->getActiveSheet()->setTitle($sheet->name);
$sheetCounter++;
}

$objPHPExcel->setActiveSheetIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');

$objWriter->save($filePath);


// echo file_get_contents($filePath);