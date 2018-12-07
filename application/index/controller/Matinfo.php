<?php
  
namespace app\index\controller;

use think\Controller;
use think\Request;
use think\Db;
use think\Loader;
use PHPExcel_IOFactory;
use PHPExcel;
class Matinfo extends Controller
{
	/**
    **导出Excel
    **
    **/
    public function export(){
        
        
        $dir=dirname(__FILE__); 
        //导出excel实现 、
        $objPHPExcel = new PHPExcel(); //创建excel对象
        $objSheet = $objPHPExcel->getActiveSheet(); //获取活动sheet
        //第一行数据 
        $objSheet->setTitle('tp5集成PHPExcel');
        $objSheet->setCellValue('A1','代码');
        $objSheet->setCellValue('B1','名称');
        $objSheet->setCellValue('C1','规格');
        $objSheet->setCellValue('D1','单位');
        $objSheet->setCellValue('E1','系数');

        //查询数据库，填充表格    
        $sql="select   matid,matname,style,unit,zsrate from matinfo";
        $rows = Db::query($sql);
        $index = 2;
        foreach ($rows as $key => $val) {
            $objSheet->setCellValue('A'.$index, $val['matid']);
            $objSheet->setCellValue('B'.$index, $val['matname']);
            $objSheet->setCellValue('C'.$index, $val['style']);
            $objSheet->setCellValue('D'.$index, $val['unit']);
            $objSheet->setCellValue('F'.$index, $val['zsrate']);
            $index++;
        }
        // var_dump($rows);
        // exit;

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5'); //生成Excel文件
		// $objWriter->save("D:/export_1.xls");//保存文件

        self::browser_export('Excel5', "browser_excel03.xls");//输出文件到浏览器
        $objWriter->save('php://output');

    }

    /**
    **
    ** 保存Excel到浏览器
    **/
    function browser_export($type, $filename){
        ob_end_clean();//清除缓冲区,避免乱码
        if($type == 'Excel5'){
            header('Content-Type: application/vnd.ms-excel'); //告诉浏览器将要输出Excel03文件
        }else{
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');//告诉浏览器将要输出Excel07文件
        }
        header('Content-Disposition: attachment;filename="'.$filename.'"');//告诉浏览器输出文件的名字
        header('Cache-Control: max-age=0');//禁止缓存
    }




    /**
    **根据下标获取单元格所在列位置
    **/
   function getCells($index){
        $arr = range('A', 'Z');
        //$arr = (A,B,C,D,E,F,G,H,I,J,K,L,M,N....Z);
        return $arr[$index]; 
    }

    /**
    **获取不同的边框样式
    **/
    function getBorderStyle($color){
        $styleArray = array(
            'borders' => array(
                'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THICK,
                    'color' => array('argb' => $color),
                ),
            ),
        );
        return $styleArray;
    } 
}