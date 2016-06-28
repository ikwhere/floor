<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dbconn.asp"-->
<!--#include file="upload_5xsoft.inc"-->
<%
dim picdiscribe,addr,contract,email,tel
set upload=new upload_5xsoft
set file=upload.file("mefile")
picdiscribe=upload.form("picdiscribe")
addr=upload.form("addr")
contract=upload.form("contract")
email=upload.form("email")
tel=upload.form("tel")
if file.fileSize>0 then

 filetype=right(file.FileName,4)
 filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&filetype
 file.saveAs Server.mappath(("..\Flags\")&filename)
 	set dbcmd=server.CreateObject("adodb.Command")
	bigpicdescribe=upload.form("kind")
	strsql="Insert into picpath(picpath,picdiscribe,addr,contract,email,tel) values ('"&filename&"','"&picdiscribe&"','"&addr&"','"&contract&"','"&email&"','"&tel&"')"
	dbcmd.ActiveConnection=db
	dbcmd.CommandText=strsql
	dbcmd.Execute
end if
set file=nothing
set upload=nothing
%>
<div style="left:100px; top:100px">您提交的信息</div>
<div style="left:100px; top:150px">联系人：<%=contract%></div>
<div style="left:100px; top:200px">厂家地址:<%=addr%></div>
<div style="left:100px; top:200px">电子邮箱:<%=email%></div>
<div style="left:100px; top:200px">联系方式:<%=tel%></div>
<div style="left:100px; top:200px">图片描述:<%=picdiscribe%></div>
<div style="left:100px; top:250px">已成功上传,<a href='uploadbigpic.asp'>点击返回</a></div>
<div style="left:100px; top:250px">
<img src="<%="../Flags/"&filename%>"></img>
</div>
</body> 
</HTML> 
