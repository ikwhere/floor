<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html> 
<body> 
<center> 
	 <form name="mainForm" enctype="multipart/form-data" action="process.asp" method=post> 
　
  <table width="70%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="28"><div align="center">效果图维护界面</div></td>
    </tr>

    <tr>
      <td height="37">请选择分类
        <select name="kind" id="kind">
        <option value="卧室">卧室</option>
        <option value="客厅">客厅</option>
        <option value="走廊">走廊</option>
        <option value="卫生间">卫生间</option>
      </select></td>
    </tr>


    <tr>
      <td height="85" valign="middle"><p>请点击“浏览”选择图片</p>
        <p>            
            <input type="file" name="mefile" />
            <br>
        　　 </p></td>
    </tr>
    <tr>
      <td height="38"><div align="center">
        <input type="submit" name="ok" value="提交" />
      </div>
   
     </td>
    </tr>
  </table>
     </form> 
  <p align="left">&nbsp;    </p>

</center> 
</body> 
</html> 
