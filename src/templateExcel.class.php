<?php 

namespace Xsp\templateExcel;

class templateExcel
{
	private $activeSheet;
	private $objPHPExcel;
	function __construct($template_path)
	{

		$objReader = \PHPExcel_IOFactory::createReader('Excel5');
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
	public function functionAutoFill($startPos,$endPos){


		$_startRe=self::getCellRowCol($startPos);
		$_endRe=self::getCellRowCol($endPos);
		$_startRe['col']=self::ordCol($_startRe['col']);
		$_endRe['col']=self::ordCol($_endRe['col']);
		
		$colLen=$_endRe['col']-$_startRe['col'];
		$rowLen=$_endRe['row']-$_startRe['row'];
		$colSign=($colLen>0)?1:-1;
		$rowSign=($rowLen>0)?1:-1;
		$colLast=($colLen>0)?0:1;
		$rowLast=($rowLen>0)?0:1;
		$colLen=abs($colLen);
		$rowLen=abs($rowLen);

		$activeSheet=$this->getActiveSheet();
		
		$colFunStr= $activeSheet->getCell($startPos)->getValue();
		for ($i=0; $i <=$colLen; $i++) { 

			if($i!=0) $colFunStr=self::getAutoFunStr($colFunStr,$colLast,1);
			$rowFunStr=$colFunStr;
			for ($j=0; $j <=$rowLen; $j++) { 
				if($j!=0) $rowFunStr=self::getAutoFunStr($rowFunStr,$rowLast,0);
				$activeSheet->setCellValueByColumnAndRow( $_startRe['col']+$colSign*$i ,$_startRe['row']+$rowSign*$j,$rowFunStr);

			}

		}


	}
	public function save($save_path)
	{
		$objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');
		var_dump($save_path);
		$objWriter->save($save_path);
	}
	/**
	*$orientation:	0:行,1:列;
	*
	*/
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
			return array('row'=>$re[2],'col'=>$re[1]);
		}else{
			return false;
		}

	}
	// 对函数字符串进行改变
	static function getAboveAutoFunStr($funStr){
		return self::getAutoFunStr($funStr,1,0);
	}
	static function getBottomAutoFunStr($funStr){
		return self::getAutoFunStr($funStr,0,0);
	}
	static function getLeftAutoFunStr($funStr){
		return self::getAutoFunStr($funStr,1,1);
	}
	static function getRightAutoFunStr($funStr){
		return self::getAutoFunStr($funStr,0,1);
	}

	// $last 		0:next 1:last
	// $orientation 	0:Row 1:Col
	static function getAutoFunStr($funStr,  $last=0, $orientation=0)
	{
		if($last==1){
			$replace_callback=($orientation==1)?(function ($m) {
				return self::getLastCol($m[0]);
			}):(function ($m) {
				var_dump($m[0]);
				return self::getLastRow($m[0]);
			});
			
		}else{
			$replace_callback=($orientation==1)?(function ($m) {
				return self::getNextCol($m[0]);
			}):(function ($m) {
				return self::getNextRow($m[0]);
			});

		}

		$pattern=array("/(?<=[A-Za-z])\d+/","/(?<!\\$)[A-Za-z]+(?=\\$?\d+)/");
		$re = preg_replace_callback(
			$pattern[$orientation], 
			$replace_callback,
			$funStr);
		return $re;
	}
	
	static function getLastRow($vaule){
		return --$vaule;
	}
	static function getNextRow($vaule){
		return ++$vaule;
	}
	/**
	* 获得上一列
	* getLastCol('A') : ''
	* getLastCol('B') : 'A'
	* getLastCol('AA') : 'Z'
	* @parma char 当前列号
	*/
	static function getLastCol($vaule){
		$vaule=strtoupper($vaule);
		$len=strlen($vaule);
		for ($i=$len-1; $i >=0 ; $i--) { 
			$v=ord($vaule[$i])-1;
			if($v<65){	//进一字符变为A
				if($i==0){
					$vaule=substr($vaule,1);
					break;
				}else{
					$vaule[$i]='Z';
				}
			}else{
				$vaule[$i]=chr($v);
				break;	//如果没有进一就结束循环
			}
		}
		return $vaule;
	}

	/**
	* 获得下一列
	* getNextCol('A') : 'B'
	* getNextCol('Z') : 'AA'
	* getNextCol('x') : 'Y'
	* @parma char 当前列号
	*/
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

	/**
	* 列号转化为列序号
	* ordCol('A') : 1
	* ordCol('b') : '2'
	* ordCol('AA') : 27
	* @parma char 列号
	*/
	static function ordCol($col){
		$col=strtoupper($col);
		$colNum=0;
		$len=strlen($col);
		for ($i=0; $i <$len ; $i++) { 
			$colNum+=(ord($col[$i])-64)*pow(26, $len-$i-1);
		}
		return $colNum-1; // form 0
	}
	/**
	* 列序号转化为列号
	* chrCol(0) : 'A'
	* chrCol(26) : 'AA'
	* @parma int  列序号
	*/
	static function chrCol($colNum){
		return ($colNum>25?self::chrCol($colNum / 26-1):'').chr($colNum % 26+65);
	}
}
?>