<?php

namespace app\index\controller;

use think\Controller;
use think\Request;
use think\Db;
use app\index\model\Matin as MatinMOdel;

class Matin extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
        $list = MatinModel::paginate(3);
        return json($list);
    }

    /**
     * 显示创建资源表单页.
     *
     * @return \think\Response
     */
    public function create()
    {
        //材料入库页面
    }

    /**
     * 保存新建的资源
     *
     * @param  \think\Request  $request
     * @return \think\Response
     */
    public function save(Request $request)
    {
        //新增材料入库记录
        $data   = $request->param();
        $result = MatinModel::create($data); 
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
        $result = MatinModel::get($id);
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
        $result = MatinModel::update($data,['ODDID' => $id]);
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
        $result = MatinModel::destroy($id);
        return json($result);
    }

    /**
    *按材料汇总
    *
    */
    public function byMatname($pdateS,$pdateE){

        $sql="select  MatType,MatName,sum(WNet) as Quality,count(1) as carNum from matin 
        where pdate between  '".$pdateS."' and  '".$pdateE."' group by MatType,MatName";
        // return $sql;
        $rows = Db::query($sql);
        return json_encode($rows);
    }


    /**
    *按供应商汇总
    *
    */
    public function bySupplier($pdateS,$pdateE){
        $sql="select  SupplierName,MatType,MatName,sum(WNet) as Quality,count(1) as carNum from matin 
            where pdate between  '".$pdateS."' and  '".$pdateE."' group by SupplierName,MatType,MatName";
        $rows = Db::query($sql);
        return json_encode($rows);
    }
}
