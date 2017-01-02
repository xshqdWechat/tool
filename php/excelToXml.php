<?php
/**
 * Created by PhpStorm.
 * User: huangjian
 * Date: 2016/12/29
 * Time: 16:53
 */
require_once '../lib/php/Classes/PHPExcel.php';

class Conversion
{
    private $file;
    private $objPHPExcel;
    private $objReader;
    private $XMLFileName = [];
    private $xmlData;
    private $xmlReader;
    private $objWriter;
    private $XMLTmpArray = [];

//    EXCEL转换为XML
    public function excelToXml($path)
    {
        $this->loadExcelFile($path);
        $this->xmlData = $this->GetExcelToXmlData();
        $this->saveToXml($this->xmlData);

    }

//    XML转换为EXCEL
    public function XmlToExcel($file)
    {
        $this->objPHPExcel = new PHPExcel();
        $this->xmlReader = simplexml_load_file($file);
        $nameArray = array();
//        获取属性名字
        foreach ($this->xmlReader->children()->children() as $name=>$txt){
            $nameArray[] = $name;
        }
        $this->XMLTmpArray[] = $nameArray;
//        遍历属性值
        foreach ($this->xmlReader->item as $key => $value) {
            $tmpArray =array();
            foreach ($value->children() as $$key => $$value) {
                $tmpArray[] = $$value;
            }
            $this->XMLTmpArray[] = $tmpArray;
        }
        $this->objPHPExcel->getActiveSheet()
                            ->fromArray($this->XMLTmpArray,null,'A1');
        $this->objWriter = $this->getPHPExcelWriter();
        $this->objWriter->save(str_replace('.php', '.xlsx', __FILE__ ));
    }


//  获取PHPExcel写对象
    public function getPHPExcelWriter()
    {
        return PHPExcel_IOFactory::createWriter($this->objPHPExcel,'Excel2007');
    }

//    获取读取对象
    public function getPHPExcelReader()
    {
        return PHPExcel_IOFactory::createReader('Excel2007');
    }

//    加载EXCEL文件
    public function loadExcelFile($path)
    {
        $this->file = $path;
        $this->objReader = $this->getPHPExcelReader();
        $this->objReader->setReadDataOnly(TRUE);
        $this->objPHPExcel = $this->objReader->load($path);
    }

//    读取Excel数据并转换为XML
    public function GetExcelToXmlData()
    {
        $objWorkSheet = $this->objPHPExcel->getActiveSheet();

        $this->getXMLFileName($objWorkSheet);
        $xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";
        $xml .= "<root>\n";

        foreach ($objWorkSheet->getRowIterator() as $rowKey => $row) {
            if ($rowKey == 1) continue;
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(FALSE);
            $xml .= "<item>\n";
            foreach ($cellIterator as $key => $cell) {
                $xml .= $this->createXmlItem($key, $cell);
            }
            $xml .= "</item>\n";
        }
        $xml .= "</root>\n";
        return $xml;

    }

//    获取XML字段
    public function getXMLFileName($workSheet)
    {
        $highestColumn = $workSheet->getHighestColumn();
        for ($row = 1; $row < 2; $row++) {
            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $this->XMLFileName[] = $workSheet->getCell($col . $row)->getValue();
            }
        }
    }

//    创建XML单项
    public function createXmlItem($key, $value)
    {
        $item = "<" . $this->XMLFileName[$key] . ">";
        $item .= $value;
        $item .= "</" . $this->XMLFileName[$key] . ">\n";
        return $item;
    }

//    将数据保存到XML文本里
    public function saveToXml($txt)
    {
        $XMLFile = fopen('xmlfile.xml', 'w') or die('无法创建/打开文件');
        fwrite($XMLFile, $txt);
        fclose($XMLFile);
    }

//    下载XML文件
    public function downXml($file = 'xmlfile.xml')
    {
        if (file_exists($file)) {
            header('Content-Type: application/xml;');
            header('Content-Disposition: attachment;filename="' . $file . '"');
            header('Content-Length: ' . filesize($file));
            header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header('Pragma: public'); // HTTP/1.0
            readfile($file);
        }

    }
//    下载EXCEL文件
    public function downExcel($file='excelToXml.xlsx'){
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$file.'"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
        $this->objWriter->save('php://output');
    }
}

$excelToXml = new Conversion();
$excelToXml->excelToXml('EN-20161223.xlsx');
//$excelToXml->downXml();

$xmlToExcel = new Conversion();
$xmlToExcel->XmlToExcel("xmlfile.xml");
$xmlToExcel->downExcel();

?>