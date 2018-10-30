<?php

namespace app\index\controller;
use app\index\model\Mbetoninfo as MbetoninfoModel;
use think\Controller;
use think\Request;
use think\Db;

class Mbetoninfo extends Controller
{
    /**
     *混凝土产品 查询强度等级以及特殊要求
     *
     * @return \think\Response
     *@param type 产品类型
     *@param str 产品名称
     */
    public function index($type,$str='')
    {
        //判断type,TSName:特殊要求，Grade:强度等级
        if($type=='grade'){
            $sql="select jsid,grade,tsid,tsname from Mbetoninfo where grade like '%".$str."%' and grade<>'' order by grade";
        }else if($type=='tsname'){
            $sql="select jsid,grade,tsid,tsname from Mbetoninfo where tsname like '%".$str."%' and tsname<>'' and grade=''";
        }

        // return $sql;
        $rows = Db::query($sql);
        return json_encode($rows);

        //显示所有产品
        // $result = MbetoninfoModel::select();
        // return json_encode($result); 
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
        $result = SyhqxdModel::create($data);
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
       $result = SyhqxdModel::get($id); 
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
        $result = SyhqxdModel::update($data,['BMID' =>$id]);
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
        $result = SyhqxdModel::destroy($id);
        return json($result);
    }

}
