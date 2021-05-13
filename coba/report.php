<?php
	require 'vendor/autoload.php';//open library
	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

	$spreadsheet = new Spreadsheet();//membuat objek dari konstruktor
	$sheet = $spreadsheet->getActiveSheet();
	$sheet->setCellValue('A1','Hello World!');

	$writer = new Xlsx($spreadsheet);
	$writer -> save('hello world.xlsx');
?>