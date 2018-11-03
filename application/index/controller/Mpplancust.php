<?php

namespace app\index\controller;

use think\Controller;
use think\Request;
use think\Db;
use app\index\model\Mpplancust as MpplancustModel;

class Mpplancust extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
       // $result =  MpplancustModel::select();
       // return json_encode($result);
        $sql="select top 20 * from Mpplancust order by iden_id desc";
        $rows = Db::query($sql);
        return json_encode($rows);
    }

    /**
     * 显示创建资源表单页.
     *
     * @return \think\Response
     */
    public function create()
    {
        
    }

    /**
     * 保存新建的资源
     *
     * @param  \think\Request  $request
     * @return \think\Response
     */
    public function save(Request $request)
    {
        $data   = $request->param();
        $result = MpplancustModel::create($data);
        return json_encode($result);
    }

    /**
     * 显示指定的资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function read($id)
    {
       $result = MpplancustModel::get($id);
       return json_encode($result);
    }

    /**
     * 显示编辑资源表单页.
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function edit($id)
    {
        //
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
       $result = MpplancustModel::update($data,['Planid' => $id]);
       return json($result);
    }

    /**
     * 删除指定资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function delete($id)
    {

        $result = MpplancustModel::destroy($id);

        return json($result);
    }
}
