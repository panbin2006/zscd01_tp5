<?php
  
namespace app\index\controller;

use think\Controller;
use think\Request;
use think\Db;
use think\Loader;
use PHPExcel_IOFactory;
use PHPExcel;
class Msaleday extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
       $sql="select custname,projectname,part,(Grade+TSName) grade,btrans,quality,remark1 from MSaleDay ";
       $rows = Db::query($sql);
       return $rows;
    }


	/**
    **导出Excel
    **
    **/
    public function export(){
        
        
        $dir=dirname(__FILE__); 
        //导出excel实现 、
        $objPHPExcel = new PHPExcel(); //创建excel对象
        $objSheet = $objPHPExcel->getActiveSheet(); //获取活动sheet
        //设置单元格样式
        $objSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->getVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//设置单元格默认对齐方式
        $objSheet->getDefaultStyle()->getAlignment()->setWrapText(true); //设置默认自动换行
        $objSheet->getColumnDimension('B')->setWidth(30);//设置“施工单位”列宽
        $objSheet->getColumnDimension('C')->setWidth(30);//设置“工程名称”列宽
        $objSheet->getColumnDimension('D')->setWidth(30);//设置“部位”列宽
        //第一行数据 
        $objSheet->setTitle('日报表');
        $objSheet->setCellValue('D1','东莞市协同混凝土有限公司'); //报表抬头，公司名称
        $objSheet->setCellValue('D2','日 报 表'); //报表标题
        $objSheet->getStyle('D2')->getFont()->setSize(30)->setBold(true);
        $objSheet->setCellValue('B3','销售（方）');
        $objSheet->setCellValue('H3','2018年12月20日');
        $objSheet->setCellValue('D1','东莞市协同混凝土有限公司');
        $objSheet->setCellValue('A4','序号');
        $objSheet->setCellValue('B4','施工单位');
        $objSheet->setCellValue('C4','工程名称');
        $objSheet->setCellValue('D4','工程部位');
        $objSheet->setCellValue('E4','强度');
        $objSheet->setCellValue('F4','浇注方式');
        $objSheet->setCellValue('G4','数量');
        $objSheet->setCellValue('H4','备注');



        //查询数据库，填充表格    
        $sql="select custname,projectname,part,(Grade+TSName) grade,btrans,quality,remark1 from MSaleDay ";
        $rows = Db::query($sql);
        var_dump($rows);
        // exit;
        $index = 5;
        foreach ($rows as $key => $val) {
            $objSheet->setCellValue('B'.$index, $val['custname']);
            $objSheet->setCellValue('C'.$index, $val['projectname']);
            // $objSheet->getStyle('C'.$index)->getAlignment()->setWrapText(true);
            $objSheet->setCellValue('D'.$index, $val['part']);
            $objSheet->setCellValue('E'.$index, $val['grade']);
            $objSheet->setCellValue('F'.$index, $val['btrans']);
            $objSheet->setCellValue('G'.$index, $val['quality']);
            $objSheet->getStyle('G'.$index)->getNumberFormat()->setFormatCode('#,##0.0');//格式化数字
            $objSheet->getStyle('G'.$index)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); //设置方量列数值右对齐
            $objSheet->setCellValue('H'.$index, $val['remark1']);
            $index++;
        }

        //表格添加边框
        $classBorderStyle = self::getBorderStyle('ffff00'); 
        $objSheet->getStyle("A4:H".$index)->applyFromArray($classBorderStyle);
        // var_dump($rows);
        // exit;
        //$objSheet->setCellValue('E'.$index, '=SUM(E2:E16)');

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
                    'style' => \PHPExcel_Style_Border::BORDER_THICK,
                    'color' => array('argb' => $color),
                 ),
                'vertical' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THICK,
                    'color' => array('argb' => $color),
                 ),
                'horizontal' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THICK,
                    'color' => array('argb' => $color),
                 ),
            ),
        );
        return $styleArray;
    } 
}