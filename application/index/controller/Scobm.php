<?php

namespace app\index\controller;
use app\index\model\Scobm as ScobmModel;
use think\Db;
use think\Controller;
use think\Request;

class Scobm extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
        // $rows = ScobmModel::field('bmid','bmname')->find();
        $sql="select bmid,bmname from SCoBM";
        $rows = Db::query($sql);
        echo json_encode($rows);
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
        $result = ScobmModel::create($data);
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
       $result = ScobmModel::get($id); 
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
        $result = ScobmModel::update($data,['BMID' =>$id]);
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
        $result = ScobmModel::destroy($id);
        return json($result);
    }
}
