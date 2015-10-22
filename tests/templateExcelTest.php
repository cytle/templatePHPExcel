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
            array('宁波市本级',1,1.70,1,1.28,1.36,1,1.56,1.93,1,1.35,1,1.3646,1.89,1,1.99,1,1.67,1.76,1,1.12),
            array('江北区',1,1.61,1,1.06,1.54,1,1.00,1.00,1,1.40,1,1.0000,1.00,1,1.00,1,1.02,1.54,1,1.19),
            array('北仑区',1,1.09,1,1.20,1.02,1,1.47,1.77,1,1.96,1,1.9621,1.82,1,1.00,1,1.59,1.64,1,1.07),
            array('镇海区',1,1.49,1,1.94,1.01,1,1.34,1.20,1,1.85,1,1.8534,1.88,1,1.00,1,1.13,1.91,1,1.17),
            array('鄞州区',1,1.94,1,1.04,1.09,1,1.31,1.40,1,1.14,1,1.5266,1.88,1,1.61,1,1.16,1.69,1,1.33),
            array('象山县',1,1.22,1,1.46,1.62,1,1.23,1.00,1,1.23,1,1.7273,1.00,1,1.50,1,1.90,1.62,1,1.87),
            array('宁海县',1,1.30,1,1.12,1.01,1,1.16,1.09,1,1.64,1,1.1349,1.00,1,1.50,1,1.65,1.92,1,1.85),
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