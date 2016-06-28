<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dbconn.asp"-->
<!--#include file="upload_5xsoft.inc"-->
<% 
Sub Bin2File(sSource,posB,posLen,strPath) '将二进制数据写入到一个文件 
Dim sTemp 
Set sTemp = Server.CreateObject("ADODB.stream") 
sTemp.Type = 1 
sTemp.Mode = 3 
sTemp.Open 

sSource.Position = posB-1 
sSource.CopyTo sTemp,posLen 
sTemp.Position = 0 
sTemp.SaveToFile strPath,2 '无条件覆盖同名文件 
sTemp.Close 
Set sTemp = nothing 
End Sub 

Function Bin2Str(sSource,posB,posLen) '二进制数据转换为字符串，包括汉字 
Dim sTemp, strData 
Set sTemp = Server.CreateObject("ADODB.stream") 
sTemp.Type = 1 
sTemp.Mode = 3 
sTemp.Open 

sSource.Position = posB-1 '位置计数首数不一样，这个对象是对0开始的 
sSource.CopyTo sTemp,posLen 
sTemp.Position = 0 
sTemp.Type = 2 
sTemp.CharSet = "gb2312" 
strData = sTemp.ReadText 

sTemp.Close 
Set sTemp = nothing 

Bin2Str = strData 
End Function 

Function DB2SB(bStr)'双字节字符串转换成单字节字符串 
Dim i 
DB2SB = "" 
For  i=1 To Len(bStr) 
DB2SB = DB2SB & ChrB(Asc(Mid(bStr,i,1))) 
Next 
End Function 

Function SB2DB(bStr)'单字节字符串转换成双字节字符串 
Dim i 
SB2DB = "" 
For i=1 To LenB(bStr) 
SB2DB = SB2DB & Chr(ascb(midb(bStr,i,1))) 
Next 
End Function 

Function GetFileName(strPath) '获取文件名 
GetFileName = Mid(strPath,InStrRev(strPath,"\")+1) 
End Function 

'判断函数 
Function iif(cond,expr1,expr2) 
If cond Then 
iif = expr1 
Else 
iif = expr2 
End If 
End Function 
%> 

<HTML> 
<HEAD> 
<title>文字和图片同时上传 </title> 
<meta http-equiv="Content-Type" content="text/html; charSet=gb2312"> 
</HEAD> 
<body> 
<% 
Dim sSource, binData 
Dim posB, posE, posSB, posSE 
Dim vCrlf, binSub 
Dim strCaption, strFileName, strContentType, posFB, posFL, arFile 
Dim i, j 
Dim dControl 
Dim strName,strValue 

vCrlf = ChrB(13) & ChrB(10) '定义一个单字节的回车换行符 
binSub = DB2SB("--") '定义一个单字节的“--”字符串 

Set sSource = Server.CreateObject("ADODB.stream") 
sSource.Type = 1 '指定返回数据类型 adTypeBinary=1,adTypeText=2 
sSource.Mode = 3 '指定打开模式 adModeRead=1,adModeWrite=2,adModeReadWrite=3 
sSource.open
    dim   content     
    content=request.binaryread(request.totalbytes) 
sSource.Write content '获取表单所有数据 

sSource.Position = 0 
binData = sSource.Read 

posB = InStrB(binData,binSub) '查找--的位置 
posB = InStrB(posB,binData,vCrlf) + 2 '+2是加入回车换行符本身的长度 
posB = InStrB(posB,binData,DB2SB("name=""")) + 6 

Set dControl = Server.CreateObject("Scripting.Dictionary") '用来保存表单各控件的信息 

Do Until posB=6 '循环处理表单各控件的数据 
posE = InStrB(posB,binData,DB2SB("""")) 
strName = Bin2Str(sSource,posB,posE-posB) '控件名称 

posB = posE + 1 '指针移动到“"”的后面 
posE = InStrB(posB,binData,vCrlf) 

If InStrB(midb(binData,posB,posE-posB),DB2SB("filename=""")) > 0 Then '判断成功表示这是一个file域 
posB = InStrB(posB,binData,DB2SB("filename=""")) + 10 
posE = InStrB(posB,binData,DB2SB("""")) 
If posE>posB Then '上传文件 
strFileName = GetFileName(Bin2Str(sSource,posB,posE-posB)) '文件名 
posB = InStrB(posB,binData,DB2SB("Content-Type:")) + 14 
posE = InStrB(posB,binData,vCrlf) 
strContentType = Bin2Str(sSource,posB,posE-posB) '文件类型 

posB = posE + 4 '这个地方换了两行，具体参看输出的原始二进制数据 
posE = InStrB(posB,binData,binSub) 

posFB = posB 
posFL = posE-posB-1 
strValue = strFileName & "," & strContentType & "," & posFB & "," & posFL '保存文件名、文件类型、在二进制数据的起始位置和长度 
Else 
strValue = "" 
End If 
Else '其他表单控件的值 
posB = posE + 4 '这个地方换了两行，具体参看输出的原始二进制数据 
posE = InStrB(posB,binData,vCrlf) 
strValue = Bin2Str(sSource,posB,posE-posB) 
End If 


dControl.Add strName,strValue '将表单名和内容添加到集合中 

posB = posE + 2 

posB = InStrB(posB,binData,vCrlf) + 2 
posB = InStrB(posB,binData,DB2SB("name=""")) + 6 
Loop 

filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&right(strFileName,4)    

bigpicdescribe = dControl.Item("kind") 
arFile = dControl.Item("mefile") 
If arFile <> "" Then '有文件数据上传 
arFile = Split(arFile,",") '将保存的文件名、文件类型等数据分离出来 
RESPONSE.Write(arFile(3))
strFileName = arFile(0) '文件名 
strContentType = arFile(1) '文件类型 
posFB = arFile(2) '起始位置 
posFL = arFile(3) '长度 
sSource.Position = posFB-1
binData = sSource.Read(posFL) '从表单数据中获取图片的二进制数据 

    dim   alllen,upstream,upstreamend,file     
    alllen=line1+line2+line3+6     
    set   upstream=server.createobject("adodb.stream")     
    set   upstreamend=server.createobject("adodb.stream")     
    upstream.type=1     
    upstreamend.type=1     
    upstream.open     
    upstreamend.open     
    upstream.write binData     
    upstream.position=alllen  
    upstreamend.write binData 

	//response.Write(server.MapPath(("..\Flags\")&filename))  
    upstreamend.savetofile server.MapPath(("..\Flags\")&filename)   
    upstream.close     
    upstreamend.close     
    set   upstream=nothing     
    set   upstreamend=nothing     

	set dbcmd=server.CreateObject("adodb.Command")
	strsql="Insert into bigpic(bigpicpath,bigpicdiscribe) values ('"&filename&"','"&bigpicdescribe&"')"
	dbcmd.ActiveConnection=db
	dbcmd.CommandText=strsql
	dbcmd.Execute
    'response.write("已成功上传,<a href='uploadpic.asp'>点击返回</a>")      

sSource.Close 
Set sSource = nothing 
end if

%> 
<div style="left:100px; top:50px">您提交的信息</div>
<div style="left:100px; top:70px">图片分类：<%=bigpicdescribe%></div>
<div style="left:100px; top:90px">图片保存为<%=filename%></div>
<div style="left:100px; top:110px">已成功上传,<a href='uploadbigpic.asp'>点击返回</a></div>
</body> 
</HTML> 
