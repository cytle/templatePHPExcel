<?php
namespace Xsp\templateExcel;

class templateExcelTest extends \PHPUnit_Framework_TestCase
{

    public function testOrdCol()
    {
        $this->assertEquals(0,templateExcel::ordCol('A'));
    }
    public function testChrCol()
    {
        $this->assertEquals('A',templateExcel::chrCol(0));
        $ZOrd=templateExcel::ordCol('Z');
        $AZOrd=templateExcel::ordCol('AZ');
        $this->assertEquals('Z',templateExcel::chrCol($ZOrd));
        $this->assertEquals('AA',templateExcel::chrCol($ZOrd+1));
        $this->assertEquals('Y',templateExcel::chrCol($ZOrd-1));
        $this->assertEquals('BA',templateExcel::chrCol($AZOrd+1));
    }
    public function testGetLastColTest_B()
    {

        $this->assertEquals('A',templateExcel::getLastCol('B'));
    }
    public function testGetLastColTest_A()
    {

        $this->assertEquals('',templateExcel::getLastCol('A'));
    }
    public function testGetLastColTest_AA()
    {
        $this->assertEquals('Z',templateExcel::getLastCol('AA'));
    }
    public function testGetNextColTest_B()
    {
        $this->assertEquals('C',templateExcel::getNextCol('B'));
    }
    public function testGetNextColTest_Z()
    {
        $this->assertEquals('AA',templateExcel::getNextCol('Z'));
    }

    public function testGetLeftAutoFunStr()
    {
        $this->assertEquals('=sum(D12:$B1)',templateExcel::getLeftAutoFunStr('=sum(E12:$B1)'));
    }
    public function testGetRightAutoFunStr()
    {
        $this->assertEquals('=sum(F12:C$1)',templateExcel::getRightAutoFunStr('=sum(E12:B$1)'));
    }
    public function testGetAboveAutoFunStr()
    {/*
        $this->assertEquals('=sum(E11:$B0)',templateExcel::getAboveAutoFunStr('=sum(E12:$B1)'));*/
        $this->assertEquals('=sum(E$12:B$1)',templateExcel::getAboveAutoFunStr('=sum(E$12:B$1)'));
    }
    public function testGetBottomAutoFunStr()
    {
        $this->assertEquals('=sum(E13:B$1)',templateExcel::getBottomAutoFunStr('=sum(E12:B$1)'));
    }
    public function initTemplateExcel(){
        $excel = new templateExcel("tests/2015-9-21国有建设用地供应情况.xls");

        return $excel;
    }
    public function testFunctionAutoFill(){
        $excel=$this->initTemplateExcel();
        $data = array(
            array('宁波市本级',50,97.70,5,18.28,250219.36,1,7.56,33251.93,3,22.35,1,9.3646,207425.89,2,12.99,1,0.67,593.76,45,67.12),
            array('江北区',27,68.61,3,9.06,10009.54,0,0.00,0.00,1,6.40,0,0.0000,0.00,0,0.00,4,14.02,10009.54,22,48.19),
            array('北仑区',31,149.09,20,38.20,35742.02,3,8.47,9968.77,1,4.96,1,4.9621,9973.82,0,0.00,14,22.59,15103.64,13,113.07),
            array('镇海区',31,69.49,22,51.94,130871.01,5,4.34,15489.20,3,21.85,3,21.8534,79654.88,0,0.00,11,11.13,10949.91,12,32.17),
            array('鄞州区',87,218.94,62,107.04,651457.09,11,9.31,29631.40,24,87.14,10,55.5266,591824.88,14,31.61,39,41.16,29138.69,13,81.33),
            array('象山县',106,215.22,73,163.46,127334.62,29,106.23,88353.00,15,27.23,8,15.7273,26013.00,7,11.50,33,39.90,12432.62,29,41.87),
            array('宁海县',81,189.30,54,93.12,113852.01,13,30.16,65395.09,5,14.64,2,7.1349,29870.00,3,7.50,38,55.65,18451.92,25,88.85),
            );
        $excel->setCellValue('N2', 'oooooooooooooooo');

        $re=$excel->setRangeValue('B8',$data);

        $excel->getActiveSheet()->mergeCells('A8:A'.$re['endRow'])->setCellValue('A8','宁波市');

        $endRow=$re['endRow'];
        $funStartCell='V'.($endRow+1);
        $excel->setCellValue($funStartCell, '=sum(V8:V'.$endRow.')');
        $excel->functionAutoFill($funStartCell, 'C'.($endRow+1));


       $this->assertEquals('=sum(C8:C'.$endRow.')',$excel->getActiveSheet()->getCell('C'.($endRow+1))->getValue());

        $excel->save(str_replace('.php', '.xls', __FILE__));
    }
}