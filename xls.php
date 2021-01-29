<?php
ini_set('error_reporting', E_ALL);
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$header = array(
    'A1' => 'Дата',
    'B1' => 'Звонки/Письма',
    'C1' => 'ЕИ по B2B МТС',
    'D1' => 'ЕИ по B2B МГТС',
    'E1' => 'ЕИ по B2G(СЗО) МТС',
    'F1' => 'ЕИ по B2O МТС',
    'G1' => 'Эскалации',
    'H1' => 'Работа с МИ',
    'I1' => 'Доп. задачи'
);
$spreadsheet1 = new Spreadsheet();
$spreadsheet2 = new Spreadsheet();
$spreadsheet3 = new Spreadsheet();


$spreadsheet2 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet2, 'Sheet2');
$spreadsheet2->fromArray(
    $header,  // The data to set
    NULL,        // Array values with this value will not be set
    'A1'         // Top left coordinate of the worksheet range where
//    we want to set these values (default is A1)
);
$spreadsheet3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet3, 'Sheet3');
$spreadsheet3->fromArray(
    $header,  // The data to set
    NULL,        // Array values with this value will not be set
    'A1'         // Top left coordinate of the worksheet range where
//    we want to set these values (default is A1)
);
//$spreadsheet1->setTitle("Sheet1");
$spreadsheet1->addSheet($spreadsheet2, 0);
$spreadsheet1->addSheet($spreadsheet3, 0);
$sheet1 = $spreadsheet1->getActiveSheet();
$sheet1->setTitle("Sheet1");
$sheet1->fromArray(
    $header,  // The data to set
    NULL,        // Array values with this value will not be set
    'A1'         // Top left coordinate of the worksheet range where
//    we want to set these values (default is A1)
);

$writer = new Xlsx($spreadsheet1);
//$writer->save('hello world.xlsx');

$fileName = 'hello world.xlsx';
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="'. urlencode($fileName).'"');
$writer->save('php://output');
