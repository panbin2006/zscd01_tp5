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
        $objSheet->setTitle('按生产日期导出');
        $objSheet->setCellValue('D1','东莞市协同混凝土有限公司'); //报表抬头，公司名称
        $objSheet->setCellValue('D2','销售列表'); //报表标题
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
    **@param String  $pdate   要生成报表的日期 '2018-12-01'
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
        self::browser_export('Excel5', "销售日报表".$pdate.".xls");//输出文件到浏览器
        $objWriter->save('php://output');
    }



    /**
    **销售明细月报表
    **@param  String  $start  开始时间
    **@param  String  $end    截止时间
    **/
    public function monthDetail($start, $end){
        $year=substr($start, 0,4);//获取年份
        $month=substr($start,5,2);//获取月份
        $filename=$year.'年('.$month.')月销售明细';
        $dir=dirname(__FILE__); //获取当前php脚本所在路径
        //导出excel实现 、
        $objPHPExcel = new PHPExcel(); //创建excel对象
        $objSheet = $objPHPExcel->getActiveSheet(); //获取活动sheet
        //设置默认单元格样式
        $objSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//设置单元格默认对齐方式
        $objSheet->getDefaultStyle()->getAlignment()->setWrapText(true); //设置默认自动换行
        $objSheet->getDefaultStyle()->getFont()->setSize(9);//设置默认字体
        //报表抬头
        $objSheet->setTitle('销售明细表');
        $objSheet->setCellValue('C1','东莞市协同混凝土有限公司'); //报表抬头，公司名称
        $objSheet->mergeCells('C1:F1');//合并公司名称单元格
        $objSheet->setCellValue('C2','2018年十二月混凝土销售明细表'); //报表标题
        $objSheet->getStyle('C2')->getFont()->setSize(18)->setBold(true);//格式化报表标题
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
        $objSheet->getColumnDimension('A')->setWidth(7.33);//设置"序号"宽度
        $objSheet->getColumnDimension('B')->setWidth(39.67);//设置"施工单位"宽度
        $objSheet->getColumnDimension('C')->setWidth(45.17);//设置"工程名称"宽度
        $objSheet->getColumnDimension('D')->setWidth(14.83);//设置"标号"宽度
        $objSheet->getColumnDimension('E')->setWidth(15.83);//设置"浇注方式"宽度
        $objSheet->getColumnDimension('F')->setWidth(12.17);//设置"数量"宽度
        $objSheet->getColumnDimension('G')->setWidth(17.17);//设置"方数合计"宽度
        $objSheet->getColumnDimension('H')->setWidth(8.5);//设置"备注"宽度
        $objSheet->getStyle('A4:H4')->getFont()->setSize(11)->setBold(true);//设置表头字体
        $sql = "select ClassName1,custname,projectname,grade,BTrans,sum(Quality) fl from MSaleDay where pdate>='".$start."' and pdate<='".$end."'group by classname1,CustName,ProjectName,Grade,BTrans";//按业务员等汇总数据
        // echo $sql;
        $rows=Db::query($sql);
        $rows = self::formatArray($rows);//按业务员、客户、工程、强度、施工方式等生成多维数组
        //var_dump($rows);
        // exit;
        $index = 5;//填充数据首行行号
        $num = 1;  //序号
        $month_sum=[];//业务员数量合计单元格数组
        foreach ($rows as $class_key => $custs) {//按业务员遍历
            $class_start=$index; //业务员起始行号
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
            $class_end=$index-1; //业务员结束行号
            $class_export=$index; //业务员汇总行的行号
            $index++;
           /*
            *业务员合计 
            */
            $objSheet->setCellValue('H'.$class_end, $class_start.':'.$class_end);//打印开始与结束index到“备注”列
            $objSheet->setCellValue('A'.$class_export, $class_key.$month."月销售合计：");//业务员汇总行
            $objSheet->setCellValue('F'.$class_export, "=SUM(F".$class_start.":F".$class_end.")"); //业务员“数量”列汇总

            $objSheet->mergeCells("A".$class_export.":E".$class_export);//合并业务行
            $objSheet->mergeCells("F".$class_export.":G".$class_export);//合并业务员合计列
            $objSheet->getStyle("A".$class_export)->getFont()->setSize(12)->setBold(true);//格式化业务员行
            $objSheet->getStyle("F".$class_export)->getFont()->setSize(12)->setBold(true);//格式业务员“数量”列汇总

            array_push($month_sum,$class_export); //业务员合计单元格位置保存到数组
        }
        //总计
        $monthSumIndex = $index + 1;
        $objSheet->setCellValue("A".$monthSumIndex, $month."月销售总计："); //总计标题单元格
        $objSheet->mergeCells("A".$monthSumIndex.":E".$monthSumIndex); //合并总计标题单元格
        
        $monthSumStr=self::getMonthSumStr($month_sum, 'F');//返回月合计单元格求和公式
        $objSheet->setCellValue("F".$monthSumIndex, $monthSumStr);
        $objSheet->mergeCells("F".$monthSumIndex.":G".$monthSumIndex); //总计数量单元格
        $objSheet->getStyle("A".$monthSumIndex)->getFont()->setSize(12)->setBold(true);//格式化总计单元格
        $objSheet->getStyle("F".$monthSumIndex)->getFont()->setSize(12)->setBold(true);//格式化"数量"合计
        
        //输出excel文件到浏览器
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5'); //生成Excel文件
        // $objWriter->save("D:/export_1.xls");//保存文件
        self::browser_export('Excel5', $filename.".xls");//输出文件到浏览器
        $objWriter->save('php://output');
    }


    /**
    **销售对账单
    **@param    String   $satrt         开始日期
    **@param    String   $end           截止日期
    **@param    String   $custName      客户名称
    **/
    public function custDetail($start, $end, $custid){
        
        $year=substr($start, 0,4);//获取年份
        $month=substr($start,5,2);//获取月份
        $filename=$year.'年('.$month.')对账单';
        $dir=dirname(__FILE__); //获取当前php脚本所在路径
        //导出excel实现 、
        $objPHPExcel = new PHPExcel(); //创建excel对象
        $objSheet = $objPHPExcel->getActiveSheet(); //获取活动sheet
        //设置默认单元格样式
        $objSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//设置单元格默认对齐方式
        $objSheet->getDefaultStyle()->getAlignment()->setWrapText(true); //设置默认自动换行
        $objSheet->getDefaultStyle()->getFont()->setSize(9);//设置默认字体
        //报表抬头
        $objSheet->setTitle($filename);
        $objSheet->setCellValue('A1','东莞市协同混凝土有限公司'.$filename); //报表抬头，公司名称
        $objSheet->mergeCells('A1:J1');//报表抬头，公司名称
        $objSheet->getStyle('A1')->getFont()->setSize(18)->setBold(true);//格式化报表标题

        //购货单位
        $objSheet->setCellValue('A2', "购货单位：".$custid); //客户名称
        $objSheet->getStyle('A2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT); //单元格左对齐
        $objSheet->mergeCells('A2:J2'); //合并客户名称单元格
        $objSheet->getStyle('A2')->getFont()->setSize(11);
        //表头,设置列名
        $objSheet->setCellValue('A3','日期');
        $objSheet->setCellValue('B3','工程名称');
        $objSheet->setCellValue('C3','施工部位');
        $objSheet->setCellValue('D3','强度等级');
        $objSheet->setCellValue('E3','方量');
        $objSheet->setCellValue('F3','单价(元)');
        $objSheet->setCellValue('G3','泵费');
        $objSheet->setCellValue('H3','补运费');
        $objSheet->setCellValue('I3','金额(元)');
        $objSheet->setCellValue('J3','备注');
        //设置单元格宽度
        $objSheet->getColumnDimension('A')->setWidth(15);//设置"日期"宽度
        $objSheet->getColumnDimension('B')->setWidth(39.67);//设置"工程名称"宽度
        $objSheet->getColumnDimension('C')->setWidth(45.17);//设置"施工部位"宽度
        $objSheet->getColumnDimension('D')->setWidth(14.83);//设置"强度"宽度
        $objSheet->getColumnDimension('E')->setWidth(15.83);//设置"方量"宽度
        $objSheet->getColumnDimension('F')->setWidth(12.17);//设置"单价(元)"宽度
        $objSheet->getColumnDimension('G')->setWidth(17.17);//设置"泵费"宽度
        $objSheet->getColumnDimension('H')->setWidth(14.83);//设置"补运费"宽度
        $objSheet->getColumnDimension('I')->setWidth(14.83);//设置"金额(元)"宽度
        $objSheet->getColumnDimension('J')->setWidth(14.83);//设置"备注"宽度

        $objSheet->getStyle('A3:J3')->getFont()->setSize(11)->setBold(true);//设置表头字体
        
        $sql = "select convert(char(10),PDate,121) PDate,ProjectName,Part,Grade+TSName Grade,Quality,PriceTotal,MoneyBS,MoneyKZ,MoneyBQTotal,Remark1 from  MSaleDay where custid='".$custid."'  and convert(char(10),PDate,121)>='".$start."' and convert(char(10),PDate,121)<='".$end."'";//按客户汇总数据
        echo $sql;
        $rows=Db::query($sql);
        // var_dump($rows);
        // exit;

        //输出发货信息
        $index=4; //起始行号
        foreach ($rows as   $key => $val) {
            # code...
            $objSheet->setCellValue('A'.$index,$val['PDate']);//日期
            $objSheet->setCellValue('B'.$index,$val['ProjectName']);//工程名称
            $objSheet->setCellValue('C'.$index,$val['Part']);//部位
            $objSheet->setCellValue('D'.$index,$val['Grade']);//强度等级
            $objSheet->setCellValue('E'.$index,$val['Quality']);//方量
            $objSheet->setCellValue('F'.$index,$val['PriceTotal']);//单价（元）
            $objSheet->setCellValue('G'.$index,$val['MoneyBS']);//泵费
            $objSheet->setCellValue('H'.$index,$val['MoneyKZ']);//补运费
            $objSheet->setCellValue('I'.$index,"=E".$index."* F".$index."+ G".$index."+ H".$index );//金额=E1*F1+G1+H1
            $objSheet->setCellValue('J'.$index,$val['Remark1']);//备注
            $index++; //行号自增
        }

        //表头增加列过滤功能
        $objSheet->setAutoFilter('A3:J3'); //表头增加列过滤功能

        
        //金额合计
        $objSheet->setCellValue('A'.$index, "合计金额:"); //合计
        $objSheet->mergeCells("A".$index.":H".$index); //合并单元格
        $objSheet->getStyle("A".$index)->getFont()->setSize(14)->setBold(true);//合计单元格格式化
        $objSheet->setCellValue("I".$index, "=SUM(I4:I".$index.")"); //合计金额
        $objSheet->getStyle("I".$index)->getFont()->setSize(14)->setBold(true);//合计单元格格式化

        //输出excel文件到浏览器
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5'); //生成Excel文件
        self::browser_export('Excel5', $filename.".xls");//输出文件到浏览器
        $objWriter->save('php://output');
    }

    /**
    *按业务员、客户、工程、强度、施工方式等生成多维数组
    *一维数据转成多维数据
    *@param Array  $arr  传入的一维数组
    **/
    function formatArray($arr){
        //创建一个新的空数组
        $arr_new=[];
        foreach ($arr as $key => $val) {
            $arr_new[$val['ClassName1']][$val['custname']][$val['projectname']][$val['grade']][$val['BTrans']] = $val;
        }
        return $arr_new;
    }


    /**
    **生成月销售总计列的=SUM(Fxx+Fx+...)
    *@param    Array   $arr    单元格数组
    *@param    String  $column 单元格列
    *@return   String  $result  返回求和字符串
    **/
    function  getMonthSumStr($arr,$column){
        $str="=SUM(";
        foreach ($arr as $key => $val) {
            if($key==0){
                $str.=$column.$val;
            }else{
                $str.="+".$column.$val;  
            }
        }
        $result=$str.")";
        return $result;
    }

    /**
    *
    * 保存Excel到浏览器
    * @param     string  $type       excel文件类型
    * @param     string  $filename   excel文件名称
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
    **@param   $color 边框颜色
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