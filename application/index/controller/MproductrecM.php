<?php

namespace app\index\controller;

use think\Controller;
use think\Request;
use app\index\model\Mproductrecm  as MproductrecmModel;

class Mproductrecm extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
       $list =  MproductrecmModel::paginate(3);
       return json($list);
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
        $result = MproductrecmModel::create($data);
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
       $result       = MproductrecmModel::get($id);
       $productrecds = $result->mproductrecd;
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
       $data = $request->param();
       $result = MproductrecmModel::update($data,['ProductID' => $id]);
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
       $result = MproductrecmModel::destroy($id);
        // 删除该盘料的材料消耗
       // $rec = MproductrecmModel::get($id);
       // if($rec->delete()){
       //      $result = $rec->mproductrecd->deleteAll();
       // }
       return json($result);
    }
}
