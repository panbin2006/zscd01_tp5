<?php

namespace app\index\model;

use think\Model;

class Carinfo extends Model
{
	//设置主键pk
	// protected  $pk = 'carid';
	//自动过滤掉不存在的字段
	protected $field = true;
}
