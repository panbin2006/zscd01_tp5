<?php

namespace app\index\model;

use think\Model;

class Mproductrecd extends Model
{
   // 设置主键
   protected $pk    = 'ProductID';


   //自动过滤不存在的字段 
   protected $field = true;

   //定义关联关系
   protected function mproductrecm(){

   		return $this->belongsTo('Mproductrecm');
   }

}
