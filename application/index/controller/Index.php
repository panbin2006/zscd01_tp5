<?php
namespace app\index\controller;
use think\Db;
use app\index\model\Mpplan as MpplanModel;
use app\index\model\Msaleodd as MsaleoddModel;

class Index
{
    public function index($pdate='')
    {
        // 查询计划方量、已送方量、未送方量
    	$sql = "select sum(Qualityplan) As jhfl,sum(QualityGive) As ysfl, sum(QualityWS) As wsfl from mpplan where convert(char(10),pdate,121)='".$pdate."'";
        $row = Db::query($sql);

        // 查询发货车次
        $sql_count = "select COUNT(1) As carNum from Msaleodd where convert(char(10),pdate,121)='".$pdate."'";
        $row[1] = Db::query($sql_count); 

        //查询出勤车辆
        $sqlCarCount="select count(distinct(CarID)) as saleCarCount from msaleodd
            where convert(char(10),pdate,121)='".$pdate."'";
        $row[2] = Db::query($sqlCarCount);
        return json_encode($row);
    }

     public function read($pdate='')
    {
        
    }
}
