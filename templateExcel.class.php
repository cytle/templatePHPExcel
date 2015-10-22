<?php 

/** PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/../Classes/PHPExcel/IOFactory.php';
/**
* 
*/
class templateExcel
{
	private $activeSheet;
	private $objPHPExcel;
	function __construct($template_path)
	{




		echo date('H:i:s') , " Load from Excel5 template" , EOL;
		$objReader = PHPExcel_IOFactory::createReader('Excel5');
		$this->objPHPExcel = $objReader->load($template_path);


		$this->setActiveSheet();
	}
	public function setActiveSheet($sheet=null){

		return $this->activeSheet=$this->objPHPExcel->getActiveSheet($sheet);
	}
	public function getActiveSheet($sheet=null){

		return (is_null($sheet)&&!is_null($this->activeSheet))?$this->activeSheet:$this->setActiveSheet($sheet);
	}
	public function setCellValue()
	{

		return call_user_func_array(array($this->getActiveSheet(),'setCellValue'),func_get_args());
	}
	public function setRangeValue($startPos="A1",$data,$insert=true)
	{
		$RowCol=self::getCellRowCol($startPos);
		if($RowCol===false){
			return false;
		}
		if(empty($data)){
			return array('startRow'=>$RowCol['row'],'endRow'=>$RowCol['row'],'startCol'=>$RowCol['col'],'endCol'=>$RowCol['col']);
		}

		$baseRow=$RowCol['row'];
		$baseCol=ord($RowCol['col']);
		$activeSheet=$this->getActiveSheet();

		if($insert) $activeSheet->insertNewRowBefore($baseRow,count($data));
		$row = $RowCol['row'];
		$firstNextRow=true;
		foreach($data as $dataRow) {
			if($firstNextRow){
				$firstNextRow=false;
			}else{
				$row = self::getNextRow($row);
			}
			$col = $RowCol['col'];
			$firstNextCol=true;
			foreach ($dataRow as $dataItem) {
				if($firstNextCol){
					$firstNextCol=false;
				}else{
					$col = self::getNextCol($col);
				}
				
				$activeSheet->setCellValue($col.$row, $dataItem);
			}
		}
		return array('startRow'=>$RowCol['row'],'endRow'=>$row,'startCol'=>$RowCol['col'],'endCol'=>$col);

	}

	/**
	*	$orientation:	0:行,1:列;
	*/
	public function functionAutoFill($startPos,$endPos,$orientation=0){

		$orientation=$orientation==1?1:0;

		$activeSheet=$this->getActiveSheet();
		$lastCell=$startPos;
		$lastFunStr=$activeSheet->getCellValue($startPos);

		$_startRe=getCellRowCol($startPos);
		$_endRe=getCellRowCol($endPos);

		for ($i=0; $i < ; $i++) { 
			
			$lastCell=self::getNextCell($lastCell,$orientation);
			$lastFunStr=self::getNextAutoFunction($lastFunStr,$orientation);
			$activeSheet->setCellValue($lastCell, $lastFunStr);
		}

	}
	public function save($save_path)
	{
		echo date('H:i:s') , " Write to Excel5 format" , EOL;
		$objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');
		$objWriter->save($save_path);
		echo date('H:i:s') , " File written to " , str_replace('.php', '.xls', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


// Echo memory peak usage
		echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
		echo date('H:i:s') , " Done writing file" , EOL;
		echo 'File has been created in ' , getcwd() , EOL;
	}

	static function getNextCell($pos,$orientation=0){
		$re=self::getCellRowCol($pos);
		if($re){
			if($orientation==1){
				$re['col']=getNextCol($re['col']);
			}else{
				$re['row']=getNextRow($re['row']);
			}
			return $re['col'].$re['row'];
		}else{
			return false;
		}
	}
	static function getCellRowCol($pos){
		if(preg_match('/([A-Za-z]+)(\d+)/', $pos, $re)){
			return array('row'=>$re[1],'col'=>strtoupper($re[0]));
		}else{
			return false;
		}

	}
	static function getNextAutoFunction($funStr,$orientation=0)
	{
		$pattern=array("/(?<!\\$)[A-Za-z]+(?=\\$?\d)/","/(?<=[A-Za-z])\d+/");
		$re = preg_replace_callback(
			$pattern[$orientation], 
			function ($m) {
				$vaule=$m[0];
				if(is_numeric($vaule)){
					return self::getNextRow($vaule);
				}else{
					return self::getNextCol($vaule);
				}
			},
			$funStr);
		return $re;
	}

	static function getLastRow($vaule){
		return --$vaule;
	}
	static function getNextRow($vaule){
		return ++$vaule;
	}

	static function getNextCol($vaule){
		$vaule=strtoupper($vaule);
		for ($i=strlen($vaule)-1; $i >=0 ; $i--) { 
			$v=ord($vaule[$i])+1;
			if($v>90){	//进一字符变为A
				$vaule[$i]='A';
				if($i==0) $vaule.='A';
			}else{
				$vaule[$i]=chr($v);
				break;	//如果没有进一就结束循环
			}
		}
		return $vaule;
	}
	static function chrCol($colNum){
		if($colNum<0){
			
		}
		return chr($colNum % 26+65);

	}
}
?>