<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dbconn.asp"-->
<!--#include file="upload_5xsoft.inc"-->


<%
set upload=new upload_5xsoft
set file=upload.file("mefile")
bigpicdescribe=upload.form("kind")
if file.fileSize>0 then

 filetype=right(file.FileName,4)
 filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&filetype
 file.saveAs Server.mappath(("..\Flags\")&filename)
 	set dbcmd=server.CreateObject("adodb.Command")
	bigpicdescribe=upload.form("kind")
	strsql="Insert into bigpic(bigpicpath,bigpicdiscribe) values ('"&filename&"','"&bigpicdescribe&"')"
	dbcmd.ActiveConnection=db
	dbcmd.CommandText=strsql
	dbcmd.Execute
end if
set file=nothing
set upload=nothing
%>
<div style="left:100px; top:100px">���ύ����Ϣ</div>
<div style="left:100px; top:150px">ͼƬ���ࣺ<%=bigpicdescribe%></div>
<div style="left:100px; top:200px">ͼƬ����Ϊ<%=filename%></div>
<div style="left:100px; top:250px">�ѳɹ��ϴ�,<a href='uploadbigpic.asp'>�������</a></div>
<div style="left:100px; top:250px">
<img src="<%="../Flags/"&filename%>"></img>
</div>
</body> 
</HTML> 
