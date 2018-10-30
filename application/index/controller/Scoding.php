<?php

namespace app\index\controller;
use app\index\model\Scoding as ScodingModel;
use think\Controller;
use think\Request;
use think\Db;
use PDO;

class Scoding extends Controller
{
    /**
     *单据代码生成 
     *
     * @return \think\Response
     */
    public function index($ModuleID,$CoID,$IsUpdate ,$PDate ,$PType=' ')
    {  

        // thinkphp调用存储过程(所有语句要全部连接起来，一起提交执行，不然会报错)
       $sql_declare="declare  @CodeID varchar(20) ";
       $sql_code='select @CodeID as codeid;';
       $sql = $sql_declare."exec GetCodeID  '".$ModuleID."' , '".$CoID."' ,".$IsUpdate." ,'".$PDate."' ,'".$PType."' ,"." @CodeID output;".$sql_code;
       $CodeID = Db::query($sql);
       return json_encode($CodeID);
       
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
    }

    /**
     * 显示指定的资源
     *
     * @param  string  $id
     * @return \think\Response
     */
    public function read($id)
    {
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
    }

    /**
     * 删除指定资源
     *
     * @param  string  $carid
     * @return \think\Response
     */
    public function delete($id)
    {
    }

}
