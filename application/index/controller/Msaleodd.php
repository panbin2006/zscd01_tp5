<?php

namespace app\index\controller;

use think\Controller;
use think\Request;
Use think\Db;
use app\index\model\Matin as MsaleoddMOdel;

class Msaleodd extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index($pdateS,$pdateE)
    {
        //按开始时间，截止时间查询送货单

        $sql="select  * from msaleodd where pdate between  '".$pdateS."' and  '".$pdateE."' order by pdate desc";
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
        //
    }

    /**
     * 显示指定的资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function read($id)
    {
        //
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
        //
    }

    /**
     * 删除指定资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function delete($id)
    {
        //
    }
}
