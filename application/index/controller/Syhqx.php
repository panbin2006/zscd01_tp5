<?php

namespace app\index\controller;
use app\index\model\Syhqx as SyhqxModel;
use think\Controller;
use think\Request;
use think\Db;

class Syhqx extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index($bmname='')
    {
        $list = Db::name('Syhqx')->field('YHID','YHName')->where('BMName',$bmname)->select();
        return json($list);
    }


    /**
     * 保存新建的资源
     *
     * @param  \think\Request  $request
     * @return \think\Response
     */
    public function save(Request $request)
    {
        //注意表单提交时字段大小写要与数据库字段大小写一致
        $data = $request->param();
        $result = SyhqxModel::create($data);
        return json($result);
    }

    /**
     * 显示指定的资源
     *
     * @param  string  $id
     * @return \think\Response
     */
    public function read($id)
    {
       $result = SyhqxModel::get($id); 
       return json($result);
        // echo 'read';
    }

    /**
     * 显示编辑资源表单页.
     *
     * @param  string  $id
     * @return \think\Response
     */
    public function edit($id)
    {
        
    }

    /**
     * 保存更新的资源
     *
     * @param  \think\Request  $request
     * @param  int  $id
     * @return \think\Response
     */
    public function update(Request $request, $id)
    {
        $data   = $request->param();
        $result = SyhqxModel::update($data,['BMID' =>$id]);
        return json($result);
    }

    /**
     * 删除指定资源
     *
     * @param  string  $carid
     * @return \think\Response
     */
    public function delete($id)
    {
        $result = SyhqxModel::destroy($id);
        return json($result);
    }

    /**
    **用户登陆验证
    **
    */

    public function check($bmname,$yhname,$pwd){
        $rows = Db::name('Syhqx')
        ->where([
                'bmname' =>['=', $bmname],
                'yhname' =>['=', $yhname],
                'pwd'    =>['=', $pwd],
            ])->find();
           // echo (!empty($rows));
        if($rows){
            return '1';
        }else{
            return '0';
        }
    }

  



}
