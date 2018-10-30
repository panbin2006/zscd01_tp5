<?php

namespace app\index\model;

use think\Model;

class Mpactm extends Model
{
   //设置主键pk
	protected  $pk = 'ProjectID';
	//自动过滤掉不存在的字段
	protected $field = true; 

	//定义关联
	public function mpactd(){
		return $this->hasMany('Mpactd','ProjectID','ProjectID');
	}
}
