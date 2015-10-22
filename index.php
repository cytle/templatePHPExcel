<?php 
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Asia/Shanghai');
echo $html_body="=sum(AZ1:\$A108)/30,BB";
echo EOL;
echo EOL;
$pattern=array("/(?<!\\$)[A-Z]+(?=\\$?\d)/","/(?<=[A-Z])\d+/");
print_r($pattern);
echo EOL;

print_r($re);
#完成自动填充函数
function (){

}
// $re = preg_match("/\\$?[A-Z]+(?=\\$?\d)/", $html_body,$re1);

// print_r($re1);
// $patterns = array ('/(19|20)(d{2})-(d{1,2})-(d{1,2})/',
//                    '/^s*{(w+)}s*=/');
// $replace = array ('3/4/12', '$1 =');
// echo preg_replace($patterns, $replace, '{startDate} = 1999-5-27');
// print_r($html_body);
/*	
require_once dirname(__FILE__) . '/templateExcel.class.php';
$excel = new templateExcel("2015-9-21国有建设用地供应情况.xls");

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

$excel->setCellValue('N2', PHPExcel_Shared_Date::PHPToExcel(time()));

$re=$excel->setRangeValue('B8',$data);

$excel->getActiveSheet()->mergeCells('A8:A'.$re['endRow'])->setCellValue('A8','宁波市');


$endCol=ord($re['endCol']);
$startCol=ord($re['startCol']);
for ($col=$startCol; $col < $endCol; $col++) { 
	echo $col,'\n';
	$_col=chr($col);echo 
	$_col,'\n';
	$excel->setCellValue($_col.($re['endRow']+1),'=sum('.$_col.'8:'.$_col.$re['endRow'].')');
} 


$excel->save(str_replace('.php', '.xls', __FILE__));*/


?>