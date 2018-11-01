<?php

namespace app\index\controller;
use app\index\model\Bcoclass5 as Bcoclass5Model;
use think\Controller;
use think\Request;
use think\Db;

//基本资料/业务员
class Bcoclass5 extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index($bmname='')
    {
        $list = Db::name('Bcoclass1')->field('ClassID1','ClassName1')->select();
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
        $result = Bcoclass5Model::create($data);
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
       $result = Bcoclass5Model::get($id); 
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
        $result = Bcoclass5Model::update($data,['BMID' =>$id]);
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
        $result = Bcoclass5Model::destroy($id);
        return json($result);
    }


}
