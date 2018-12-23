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
    **导出日报表
    **
    **/
    public function exportDay($pdate){
        $dir=dirname(__FILE__); 
        //导出excel实现 、
        $objPHPExcel = new PHPExcel(); //创建excel对象
        $objSheet = $objPHPExcel->getActiveSheet(); //获取活动sheet
        //设置单元格样式
        $objSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//设置单元格默认对齐方式
        $objSheet->getDefaultStyle()->getAlignment()->setWrapText(true); //设置默认自动换行
        $objSheet->getColumnDimension('B')->setWidth(30);//设置“施工单位”列宽
        $objSheet->getColumnDimension('C')->setWidth(30);//设置“工程名称”列宽
        $objSheet->getColumnDimension('D')->setWidth(30);//设置“部位”列宽
        //表格抬头
        $objSheet->setTitle($pdate.'日报表');//设置sheet名称
        $objSheet->setCellValue('D1','东莞市协同混凝土有限公司'); //报表抬头，公司名称
        $objSheet->setCellValue('D2','日 报 表'); //报表标题
        $objSheet->getStyle('D2')->getFont()->setSize(30)->setBold(true);
        $objSheet->setCellValue('B3','销售（方）');
        $objSheet->setCellValue('H3',$pdate);
        $objSheet->setCellValue('D1','东莞市协同混凝土有限公司');
        //添加表头
        $objSheet->setCellValue('A4','序号');
        $objSheet->setCellValue('B4','施工单位');
        $objSheet->setCellValue('C4','工程名称');
        $objSheet->setCellValue('D4','工程部位');
        $objSheet->setCellValue('E4','强度');
        $objSheet->setCellValue('F4','浇注方式');
        $objSheet->setCellValue('G4','数量');
        $objSheet->setCellValue('H4','备注');
        $objSheet->getStyle('A4:H4')->getFont()->setSize(14)->setBold(true);//设置表头字体
        /**
        **询数据库，填充表格 
        **/
        //查询所有工程名称
        $sql="select distinct(custname) from MSaleDay where pdate='".$pdate."'"; //查询当日有生产的所有客户
        $rows = Db::query($sql);
        $index = 5;//填充数据首行行号
        $num = 1;  //序号
        foreach ($rows as $key => $val) {//循环遍历客户
            $objSheet->setCellValue('A'.$index, $num);//写入序号
            $objSheet->setCellValue('B'.$index, $val['custname']);//写入客户
            $old_index = $index; //客户生产记录首行行号
            $sql="select projectname,part,(Grade+TSName) grade,btrans,quality,remark1 from MSaleDay where pdate='".$pdate."' and CustName='".$val['custname']."'"; //查询当前客户的生产记录
            $rows = Db::query($sql);
            foreach ($rows as $key => $val) {//循环遍历客户的生产记录
                $objSheet->setCellValue('C'.$index, $val['projectname']);
                // $objSheet->getStyle('C'.$index)->getAlignment()->setWrapText(true);
                $objSheet->setCellValue('D'.$index, $val['part']);
                $objSheet->setCellValue('E'.$index, $val['grade']);
                $objSheet->setCellValue('F'.$index, $val['btrans']);
                $objSheet->setCellValue('G'.$index, $val['quality']);
                //$objSheet->getStyle('G'.$index)->getNumberFormat()->setFormatCode('#,##0.0');//格式化数字
                $objSheet->getStyle('G'.$index)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); //设置方量列数值右对齐
                $objSheet->setCellValue('H'.$index, $val['remark1']);
                $index++; //行号自增
            }
            $num++;//序号自增
            $objSheet->mergeCells('B'.$old_index.':B'.($index-1)); //合并客户名称单元格
            $objSheet->getStyle('B'.$old_index)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_TOP);//设置客户名称单元格上对齐
            $objSheet->mergeCells('A'.$old_index.':A'.($index-1)); //合并序号单元格
            $objSheet->getStyle('A'.$old_index)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_TOP);//设置序号单元格上对齐
            //var_dump('B'.$old_index.':B'.$index);
        }
        //表格添加边框
        $classBorderStyle = self::getBorderStyle('333333'); //设定边框颜色
        $objSheet->getStyle("A4:H".$index)->applyFromArray($classBorderStyle);//给表格加上边框
        //合计行导出
        $objSheet->mergeCells('B'.$index.':F'.$index); //合并合计单元格
        $objSheet->setCellValue('B'.$index, '方量合计：');
        $objSheet->setCellValue('G'.$index, '=SUM(G5:G'.($index-1).')');//方量合计数
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5'); //生成Excel文件
        // $objWriter->save("D:/export_1.xls");//保存文件
        self::browser_export('Excel5', "browser_excel03.xls");//输出文件到浏览器
        $objWriter->save('php://output');
    }



    /**
    **销售明细月报表
    **/
    public function monthDetail($start, $end){
        $year=substr($start, 0,4);//获取年份
        $month=substr($start,5,2);//获取月份
        $filename=$year.'年('.$month.')月销售明细';
        $dir=dirname(__FILE__); //获取当前php脚本所在路径
        //导出excel实现 、
        $objPHPExcel = new PHPExcel(); //创建excel对象
        $objSheet = $objPHPExcel->getActiveSheet(); //获取活动sheet
        //设置单元格样式
        $objSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//设置单元格默认对齐方式
        $objSheet->getDefaultStyle()->getAlignment()->setWrapText(true); //设置默认自动换行
        //报表抬头
        $objSheet->setTitle('销售明细表');
        $objSheet->setCellValue('C1','东莞市协同混凝土有限公司'); //报表抬头，公司名称
        $objSheet->mergeCells('C1:F1');//合并公司名称单元格
        $objSheet->setCellValue('C2','2018年十二月混凝土销售明细表'); //报表标题
        $objSheet->mergeCells('C2:F2');
        $objSheet->getStyle('D2')->getFont()->setSize(30)->setBold(true);
        $objSheet->setCellValue('B3','单位：方');
        //表头,设置列名
        $objSheet->setCellValue('A4','序号');
        $objSheet->setCellValue('B4','施工单位');
        $objSheet->setCellValue('C4','工程名称');
        $objSheet->setCellValue('D4','标号');
        $objSheet->setCellValue('E4','浇注方式');
        $objSheet->setCellValue('F4','数量');
        $objSheet->setCellValue('G4','方数合计');
        $objSheet->setCellValue('H4','备注');
        //设置单元格宽度
        $objSheet->getColumnDimension('B')->setWidth(30);//设置客户列宽度
        $objSheet->getColumnDimension('C')->setWidth(30);//设置工程列宽度
        $objSheet->getColumnDimension('G')->setWidth(10);//设置工程合计列宽度
        $objSheet->getStyle('A4:H4')->getFont()->setSize(12)->setBold(true);//设置表头字体
        $sql = "select ClassName1,custname,projectname,grade,BTrans,sum(Quality) fl from MSaleDay where pdate>='".$start."' and pdate<='".$end."'group by classname1,CustName,ProjectName,Grade,BTrans";//按业务员等汇总数据
        // echo $sql;
        $rows=Db::query($sql);
        $rows = self::formatArray($rows);//按业务员、客户、工程、强度、施工方式等生成多维数组
        var_dump($rows);
        // exit;
        $index = 5;//填充数据首行行号
        $num = 1;  //序号
        foreach ($rows as $class_key => $custs) {//按业务员遍历
            foreach ($custs as $key => $projects) {//按客户遍历
                $cust_start = $index; //当前客户起始行号
                $objSheet->setCellValue('A'.$index, $num);//客户
                foreach ($projects as $key => $grades) {//按工程遍历
                    $proj_start=$index;//当前工程起始
                    foreach ($grades as $key => $btrans) {//按强度与施工方式遍历
                        foreach ($btrans as $key => $val) {
                            // $objSheet->getStyle('C'.$index)->getAlignment()->setWrapText(true);
                            $objSheet->setCellValue('B'.$index, $val['custname']);//客户
                            $objSheet->setCellValue('C'.$index, $val['projectname']);//工程
                            $objSheet->setCellValue('D'.$index, $val['grade']);//强度
                            $objSheet->setCellValue('E'.$index, $val['BTrans']);//浇注方式
                            $objSheet->setCellValue('F'.$index, $val['fl']);//方量
                            $objSheet->getStyle('F'.$index)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); //设置方量列数值右对齐
                            //$objSheet->setCellValue('G'.$index, $index);//合计
                            // $objSheet->setCellValue('H'.$index, $val['remark1']);
                            $index++; //行号自增
                        }
                    }
                    $proj_end=$index-1;//工程结束行号
                    $objSheet->setCellValue('G'.$proj_start, '=SUM(F'.$proj_start.':F'.$proj_end.')'); //工程合计
                    $objSheet->mergeCells('C'.$proj_start.':C'.$proj_end); //合并工程列单元格
                    $objSheet->mergeCells('G'.$proj_start.':G'.$proj_end); //合并工程合计列单元格
                }

                $cust_end=$index-1;//客户结束行号
                $objSheet->mergeCells('B'.$cust_start.':B'.$cust_end); //合并客户列单元格
                $objSheet->mergeCells('A'.$cust_start.':A'.$cust_end); //合并序号列单元格
                $num++; //序号自增
            }
           //业务员合计 
            $objSheet->setCellValue('A'.$index,$class_key.'业务员合计:');//业务员合计行
        }
        //总计

        //输出excel文件到浏览器
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5'); //生成Excel文件
        // $objWriter->save("D:/export_1.xls");//保存文件
        self::browser_export('Excel5', "browser".$filename.".xls");//输出文件到浏览器
        $objWriter->save('php://output');
    }

    /**
    **按业务员、客户、工程、强度、施工方式等生成多维数组
    **一维数据转成多维数据
    **/
    function formatArray($arr){
        //创建一个新的空数组
        $arr_new=[];
        foreach ($arr as $key => $val) {
            $arr_new[$val['ClassName1']][$val['custname']][$val['projectname']][$val['grade']][$val['BTrans']] = $val;
            // foreach ($arr as $key => $val) {
            //     $arr_new[][$val['custname']] = $val;
            // }
        }

        return $arr_new;
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
                'allborders' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => $color),
                 ),
                
            ),
        );
        return $styleArray;
    } 
}