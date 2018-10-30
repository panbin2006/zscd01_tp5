<?php

namespace app\index\controller;
use app\index\model\Bpline as PlineModel;
use think\Controller;
use think\Request;

class Bpline extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
        $list = PlineModel::select();
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
        $result = PlineModel::create($data);
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
       $result = PlineModel::get($id); 
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
        $result = PlineModel::update($data,['BMID' =>$id]);
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
        $result = PlineModel::destroy($id);
        return json($result);
    }
}
