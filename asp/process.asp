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
<div style="left:100px; top:100px">���ύ����Ϣ</div>
<div style="left:100px; top:150px">��ϵ�ˣ�<%=contract%></div>
<div style="left:100px; top:200px">���ҵ�ַ:<%=addr%></div>
<div style="left:100px; top:200px">��������:<%=email%></div>
<div style="left:100px; top:200px">��ϵ��ʽ:<%=tel%></div>
<div style="left:100px; top:200px">ͼƬ����:<%=picdiscribe%></div>
<div style="left:100px; top:250px">�ѳɹ��ϴ�,<a href='uploadbigpic.asp'>�������</a></div>
<div style="left:100px; top:250px">
<img src="<%="../Flags/"&filename%>"></img>
</div>
</body> 
</HTML> 
