<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Asia/Shanghai');

/** PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/../Classes/PHPExcel/IOFactory.php';



echo date('H:i:s') , " Load from Excel5 template" , EOL;
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load("2015-9-21国有建设用地供应情况.xls");




echo date('H:i:s') , " Add new data to the template" , EOL;
$data = array(
	array('宁波市本级',50,97.70,5,18.28,250219.36,1,7.56,33251.93,3,22.35,1,9.3646,207425.89,2,12.99,1,0.67,593.76,45,67.12),
	array('江北区',27,68.61,3,9.06,10009.54,0,0.00,0.00,1,6.40,0,0.0000,0.00,0,0.00,4,14.02,10009.54,22,48.19),
	array('北仑区',31,149.09,20,38.20,35742.02,3,8.47,9968.77,1,4.96,1,4.9621,9973.82,0,0.00,14,22.59,15103.64,13,113.07),
	array('镇海区',31,69.49,22,51.94,130871.01,5,4.34,15489.20,3,21.85,3,21.8534,79654.88,0,0.00,11,11.13,10949.91,12,32.17),
	array('鄞州区',87,218.94,62,107.04,651457.09,11,9.31,29631.40,24,87.14,10,55.5266,591824.88,14,31.61,39,41.16,29138.69,13,81.33),
	array('象山县',106,215.22,73,163.46,127334.62,29,106.23,88353.00,15,27.23,8,15.7273,26013.00,7,11.50,33,39.90,12432.62,29,41.87),
	array('宁海县',81,189.30,54,93.12,113852.01,13,30.16,65395.09,5,14.64,2,7.1349,29870.00,3,7.50,38,55.65,18451.92,25,88.85),
	array('余姚市',89,155.64,59,109.32,250430.27,8,10.65,30488.85,20,39.87,9,32.1642,181427.62,11,7.71,40,65.44,37540.26,21,39.69),
	array('慈溪市',112,325.25,81,128.93,59696.61,3,12.22,10315.00,6,8.95,1,0.1685,712.00,5,8.79,76,116.18,48243.61,27,187.89),
	array('奉化市',48,137.24,28,46.20,39957.28,3,1.97,5767.00,10,19.76,2,2.4040,15041.00,8,17.35,23,41.82,19149.28,12,73.69),
	);

$objPHPExcel->getActiveSheet()->setCellValue('N2', PHPExcel_Shared_Date::PHPToExcel(time()));

$baseRow = 8;
$baseCol   = ord('B');
$activeSheet=$objPHPExcel->getActiveSheet();

$objPHPExcel->getActiveSheet()->insertNewRowBefore($baseRow,count($data));
foreach($data as $r => $dataRow) {
	$row = $baseRow + $r;
	$col=$baseCol;
	foreach ($dataRow as $key => $dataItem) {
		$activeSheet->setCellValue(chr($col++).$row, $dataItem);
	}
}


$activeSheet->mergeCells('A8:A'.$row)->setCellValue('A8','宁波市');


echo date('H:i:s') , " Write to Excel5 format" , EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xls', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
