<?php

require "Classes/PHPExcel.php";
require "Classes/PHPExcel/Writer/Excel5.php"; 

$objPHPExcel = new PHPExcel();
// Set document properties
$objPHPExcel->getActiveSheet()->setShowGridlines(false);
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

$objPHPExcel->getProperties()->setCreator("Hassan Akhlaq")
                             ->setLastModifiedBy("Hassan Akhlaq")
                             ->setTitle("First Sheet")
                             ->setSubject("Test Document")
                             ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                             ->setKeywords("office 2007 openxml php")
                             ->setCategory("Test result file");

$objPHPExcel->getActiveSheet()->getStyle('A1:Z999')->getAlignment()->setWrapText(true);
$border_bottom = array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN)));
$border_all = array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER, 'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER),
						  'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
											 'right' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
											 'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
											 'left' => array('style' => PHPExcel_Style_Border::BORDER_THICK)));
$objPHPExcel->getActiveSheet()->setCellValue("B7", "Admission # : ");
$objPHPExcel->getActiveSheet()->mergeCells("B7:C7");
$objPHPExcel->getActiveSheet()->mergeCells("B9:C9");
$objPHPExcel->getActiveSheet()->setCellValue("B9", "Roll # : ");
$objPHPExcel->getActiveSheet()->mergeCells("B12:C12");
$objPHPExcel->getActiveSheet()->setCellValue("B12", "Class / Section : ");
$objPHPExcel->getActiveSheet()->mergeCells("B14:C14");
$objPHPExcel->getActiveSheet()->setCellValue("B14", "Course : ");
$objPHPExcel->getActiveSheet()->mergeCells("B16:C16");
$objPHPExcel->getActiveSheet()->setCellValue("B16", "Session : ");

$objPHPExcel->getActiveSheet()->mergeCells("B20:D22");
$objPHPExcel->getActiveSheet()->setCellValue("B20", "SUBJECT");
$objPHPExcel->getActiveSheet()->getStyle('B20:D22')->applyFromArray($border_all);
$objPHPExcel->getActiveSheet()->mergeCells("E20:E22");
$objPHPExcel->getActiveSheet()->setCellValue("E20", "Obtain Marks");
$objPHPExcel->getActiveSheet()->getStyle('E20:E22')->applyFromArray($border_all);
$objPHPExcel->getActiveSheet()->mergeCells("F20:F22");
$objPHPExcel->getActiveSheet()->setCellValue("F20", "Total Marks");
$objPHPExcel->getActiveSheet()->getStyle('F20:F22')->applyFromArray($border_all);
$objPHPExcel->getActiveSheet()->mergeCells("G20:G22");
$objPHPExcel->getActiveSheet()->setCellValue("G20", "Status");
$objPHPExcel->getActiveSheet()->getStyle('G20:G22')->applyFromArray($border_all);
$objPHPExcel->getActiveSheet()->mergeCells("H20:I22");
$objPHPExcel->getActiveSheet()->setCellValue("H20", "Percentage");
$objPHPExcel->getActiveSheet()->getStyle('H20:I22')->applyFromArray($border_all);
$objPHPExcel->getActiveSheet()->mergeCells("J20:M22");
$objPHPExcel->getActiveSheet()->setCellValue("J20", "Instructor");
$objPHPExcel->getActiveSheet()->getStyle('J20:M22')->applyFromArray($border_all);
$x=2; $y=22;
for ($i=0; $i<10; $i++)
{
	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($x, $y+$i, "");
}

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