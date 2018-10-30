<?php

namespace app\index\model;

use think\Model;

class Mproductrecm extends Model
{
    // 设置主键
	protected $pk 	 = 'ProductID';
    // 自动过滤掉不存在的字段
    protected $field = true;

    //定义关联
    protected function mproductrecd(){
    	//关联主键字段名大小字要与数据库一致
    	return $this->hasMany('Mproductrecd','ProductID','ProductID');
    }
    
}
