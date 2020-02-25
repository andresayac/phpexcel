<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

#IMPORT LIB
require_once dirname(__FILE__) . './PHPExcel.php';


#VALUES DB-CONNECTION

define('DB_NAME', 'dbname');
define('DB_USER', 'root');
define('DB_PASSWORD', '');
define('DB_HOST', 'localhost');
define('DB_PORT', '3306');

#NAME FILE
$nameFile = "Report -".date('l \t\h\e jS');

## CONNECTION DB
$db     =   new mysqli(DB_HOST, DB_USER, DB_PASSWORD, DB_NAME,DB_PORT);

## CHECK CONNECTION DB
if ($db->connect_error) {
    die("Connection failed: " . $db->connect_error);
}

#INSTANCE PHPExcel
$objPHPExcel    =   new PHPExcel();

#QUERY
$sql            =  "SELECT * FROM country"; 
  
$result =   $db->query($sql) or die(mysql_error());

$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setTitle($nameFile);
$objPHPExcel->getProperties()->setCreator("Octopus");

//ADD ELEMENTS TO ARRAY
$datos_excel = [];

if (!$result->num_rows) exit();

while ($row = $result->fetch_assoc()) {
    array_push($datos_excel, $row);
}
array_unshift($datos_excel, array_keys($datos_excel[0]));

$objPHPExcel->getActiveSheet()->fromArray($datos_excel);

$objWriter  =   new PHPExcel_Writer_Excel2007($objPHPExcel);
  

header('Content-Type: application/vnd.ms-excel'); 
header('Content-Disposition: attachment;filename="'.$nameFile.'.xlsx"'); 
header('Cache-Control: max-age=0'); 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');  
$objWriter->save('php://output');


