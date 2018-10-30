<?php

namespace app\index\model;

use think\Model;

class Msaleodd extends Model
{
   //设置主键pk
	protected  $pk = 'SaleID';
	//自动过滤掉不存在的字段
	protected $field = true; 

	//定义关联
	public function mpplan(){
		return $this->belongsTo('Mpplan');
	}  
}
