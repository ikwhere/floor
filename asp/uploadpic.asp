<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html> 
<body> 
<center> 
	 <form name="mainForm" enctype="multipart/form-data" action="process.asp" method=post> 
　
  <table width="49%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="28"><div align="center">地板图维护界面</div></td>
    </tr>

    <tr>
      <td height="37"><div align="left">请输入联系人&nbsp;&nbsp;
            <input name="contract" type="text" id="contract" />
      </div></td>
    </tr>


    <tr>
      <td height="34" valign="middle"><p align="left">请输入厂家地址 
        <input name="addr" type="text" id="addr" />        　　 </p>        </td>
    </tr>
    <tr>
      <td height="32" valign="middle"><p align="left">请输入联系方式 
        <input name="tel" type="text" id="tel" />
      </p>        </td>
    </tr>
    <tr>
      <td height="38" valign="middle"><div align="left">请输入电子邮箱 
            <input name="email" type="text" id="email" />
      </div></td>
    </tr>
    <tr>
      <td height="69" valign="middle"><p align="left">请点击“浏览”选择图片<br><br>
          
       &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="file" name="mefile" />
        </p>
        </td>
    </tr>
    <tr>
      <td height="163" valign="middle"><p align="left">请输入图片描述</p>
        <p align="left">            <textarea name="picdiscribe" cols="50" rows="8" id="picdiscribe"></textarea>
        </p></td>
    </tr>
    <tr>
      <td height="38">
          <div align="center">
            <input type="submit" name="ok1" value="提交" />
            </div></td></tr>
  </table>
     </form> 
  <p align="left">&nbsp;    </p>

</center> 
</body> 
</html> 
