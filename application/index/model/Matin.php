<?php

namespace app\index\model;

use think\Model;

class Matin extends Model
{
	//设置主键pk
	protected  $pk = 'ODDID';
	//自动过滤掉不存在的字段
	protected $field = true;
}
