<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel.php';

// $objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:C1');

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
							
$filePath  = exist(@$json->path, 'file.xlsx');
							
$sheets  = exist(@$json->sheets, []);
unlink($file);


$sheetCounter = 0;
foreach($sheets as $sheet){

	
$datas  = exist(@$sheet->rows, []);
$colWidth = exist(@$sheet->columns, []);
$drawings = exist(@$sheet->drawings, []);

//if not default sheet create a new one
if($sheetCounter !=0 ){
	$objPHPExcel->createSheet();
}
$activeSheet = $objPHPExcel->setActiveSheetIndex($sheetCounter);

foreach($colWidth as $col){
	if(isset($col->width)){
		$activeSheet->getColumnDimension($col->key)->setWidth($col->width);
	}
}

foreach($datas as $row){
    foreach($row as $data){
		$activeSheet ->setCellValue($data->key, $data->value);
		if(isset($data->style)){
			
			//cell's type
			if(isset($data->style->format)){
				$activeSheet->getStyle($data->key) ->getNumberFormat()->setFormatCode($data->style->format);
			}

			//cell's font size
			if(isset($data->style->fontSize)){
				$activeSheet->getStyle($data->key) ->getFont()->setSize($data->style->fontSize);
			}
			//cell's bgcolor (a RGB color)
			if(isset($data->style->fill)){
				$activeSheet->getStyle($data->key) ->getFill()
				->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				->getStartColor()
				->setRGB($data->style->fill);
			}
			//cell border
			if(isset($data->style->border)){
				$style = $data->style->border->style;
				$excel_boder_style = '';
				switch ($style) {
					case 'thin':
					$excel_boder_style = PHPExcel_Style_Border::BORDER_THIN;
						break;
					
					default:
					$excel_boder_style = PHPExcel_Style_Border::BORDER_THICK;
						break;
				}
				$color = $data->style->border->color;
				$position =  @$data->style->border->position;
				$activeSheet->getStyle($data->key)->applyFromArray(
					$border_style= array('borders' => array(exist($position, 'allborders') => array('style' => 
					$excel_boder_style,'color' => array('argb' => $color),)))
				);
			}

			//font
			if(isset($data->style->font)) {
				$styleArray = array('font'  => json_decode(json_encode($data->style->font), True));
				$activeSheet->getStyle($data->key)->applyFromArray($styleArray);
			}
			//alignment and rotation
			if(isset($data->style->alignment)){
				$key = $data->style->alignment->key;
				$value = $data->style->alignment->value;
				$rotation = @$data->style->alignment->rotation;
				$styleArray = array(
					'alignment'=>array(
						$key => $value
					)
				);
				$activeSheet->getStyle($data->key)->applyFromArray($styleArray);
				$activeSheet->getStyle($data->key)->getAlignment()->setTextRotation( exist($rotation, 0));
			}
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


echo file_get_contents($filePath);