<?php
// +----------------------------------------------------------------------
// | ThinkPHP [ WE CAN DO IT JUST THINK ]
// +----------------------------------------------------------------------
// | Copyright (c) 2006~2018 http://thinkphp.cn All rights reserved.
// +----------------------------------------------------------------------
// | Licensed ( http://www.apache.org/licenses/LICENSE-2.0 )
// +----------------------------------------------------------------------
// | Author: liu21st <liu21st@gmail.com>
// +----------------------------------------------------------------------
use think\Route;
// Route::resource('carinfo', 'index/carinfo');
// //Route::resource('index', 'index/index');
// Route::resource('matin', 'index/matin');
// Route::resource('mpactm', 'index/mpactm');
// Route::resource('mpplan', 'index/mpplan');
// Route::resource('msaleodd', 'index/msaleodd');
// Route::resource('mpplancust', 'index/mpplancust');
// Route::resource('mproductrecm', 'index/mproductrecm');
// Route::resource('mproductrecd', 'index/mproductrecd');


return [
    '__pattern__' => [
        'name' => '\w+',
    ],
    '[hello]'     => [
        ':id'   => ['index/hello', ['method' => 'get'], ['id' => '\d+']],
        ':name' => ['index/hello', ['method' => 'post']],
    ],

];
