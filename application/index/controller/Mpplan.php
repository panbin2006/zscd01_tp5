<?php

namespace app\index\controller;

use think\Controller;
use think\Request;
use think\Db;
use app\index\model\Mpplan as MpplanModel;
class Mpplan extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index($pdate)
    {
        $sql="select   * from mpplan where  pdate='".$pdate."' order by planid desc";
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
        //
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
        $result = MpplanModel::create($data);
        return json($result);
    }

    /**
     * 显示指定的资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function read($id)
    {
        $result    = mpplanmodel::get($id);
        $msaleodds = $result->msaleodd;
        return json($result);
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
        $result = MpplanModel::update($data, ['planid'=>$id]);
        return json($data);
    }

    /**
     * 删除指定资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function delete($id)
    {
        $result = MpplanModel::destroy($id);
        return json($result);
    }
}
