<?php

require "Classes/PHPExcel.php";
require "Classes/PHPExcel/Writer/Excel5.php"; 

$objPHPExcel = new PHPExcel();
// Set document properties

$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

$objPHPExcel->getProperties()->setCreator("Hassan Akhlaq")
                             ->setLastModifiedBy("Hassan Akhlaq")
                             ->setTitle("First Sheet")
                             ->setSubject("Test Document")
                             ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                             ->setKeywords("office 2007 openxml php")
                             ->setCategory("Test result file");

//Add some data
$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(1);
$objPHPExcel->getActiveSheet()->setCellValue("A1", "Hassan Akhlaq");
$objPHPExcel->getActiveSheet()->mergeCells("A1:T1");
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setSize(22);
$objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("A1")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$styleArray = array(
    'borders' => array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THICK,
            'color' => array('argb' => 'FFFF0000'),
        ),
    ),
);
$sBorderStyle = array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT),
						  'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											 'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											 'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN),
											 'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN)));



$objPHPExcel->getActiveSheet()->setCellValue("B3", "");
$objPHPExcel->getActiveSheet()->mergeCells("B3:T6");


$objPHPExcel->getActiveSheet()->getStyle('B12')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLACK);

$objPHPExcel->getActiveSheet()->getStyle('B12')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objPHPExcel->getActiveSheet()->getStyle('B12')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
$objPHPExcel->getActiveSheet()->getStyle('B12')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
$objPHPExcel->getActiveSheet()->getStyle('B12')->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
$objPHPExcel->getActiveSheet()->getStyle('B12')->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->getActiveSheet()->getStyle('B12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objPHPExcel->getActiveSheet()->getStyle('B12')->getFill()->getStartColor()->setARGB('FFFF0000');

$styleArr = array(
	'font' => array(
		'bold' => true,
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
	),
	'borders' => array(
		'bottom' => array(
			'style' => PHPExcel_Style_Border::BORDER_THICK,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
		'rotation' => 90,
		'startcolor' => array(
			'argb' => 'FFA0A0A0',
		),
		'endcolor' => array(
			'argb' => 'FFFFFFFF',
		),
	),
);


$objPHPExcel->getActiveSheet()->getStyle('A1:T1')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->getStyle('H13')->applyFromArray($styleArr);
$objPHPExcel->getActiveSheet()->mergeCells("C15:P15");
$objPHPExcel->getActiveSheet()->getStyle('C15:P15')->applyFromArray($sBorderStyle);

// $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(25);
// $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
// $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);
// $objPHPExcel->setActiveSheetIndex(0)
//             ->setCellValue('A1', 'Sr no')
//             ->setCellValue('B1', 'Name')
//             ->setCellValue('C1', 'Age');
            


// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('My Sheet');
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
header('Content-Disposition: attachment;filename="userList.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
?>