<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dbconn.asp"-->
<!--#include file="upload_5xsoft.inc"-->
<% 
Sub Bin2File(sSource,posB,posLen,strPath) '������������д�뵽һ���ļ� 
Dim sTemp 
Set sTemp = Server.CreateObject("ADODB.stream") 
sTemp.Type = 1 
sTemp.Mode = 3 
sTemp.Open 

sSource.Position = posB-1 
sSource.CopyTo sTemp,posLen 
sTemp.Position = 0 
sTemp.SaveToFile strPath,2 '����������ͬ���ļ� 
sTemp.Close 
Set sTemp = nothing 
End Sub 

Function Bin2Str(sSource,posB,posLen) '����������ת��Ϊ�ַ������������� 
Dim sTemp, strData 
Set sTemp = Server.CreateObject("ADODB.stream") 
sTemp.Type = 1 
sTemp.Mode = 3 
sTemp.Open 

sSource.Position = posB-1 'λ�ü���������һ������������Ƕ�0��ʼ�� 
sSource.CopyTo sTemp,posLen 
sTemp.Position = 0 
sTemp.Type = 2 
sTemp.CharSet = "gb2312" 
strData = sTemp.ReadText 

sTemp.Close 
Set sTemp = nothing 

Bin2Str = strData 
End Function 

Function DB2SB(bStr)'˫�ֽ��ַ���ת���ɵ��ֽ��ַ��� 
Dim i 
DB2SB = "" 
For  i=1 To Len(bStr) 
DB2SB = DB2SB & ChrB(Asc(Mid(bStr,i,1))) 
Next 
End Function 

Function SB2DB(bStr)'���ֽ��ַ���ת����˫�ֽ��ַ��� 
Dim i 
SB2DB = "" 
For i=1 To LenB(bStr) 
SB2DB = SB2DB & Chr(ascb(midb(bStr,i,1))) 
Next 
End Function 

Function GetFileName(strPath) '��ȡ�ļ��� 
GetFileName = Mid(strPath,InStrRev(strPath,"\")+1) 
End Function 

'�жϺ��� 
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
<title>���ֺ�ͼƬͬʱ�ϴ� </title> 
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

vCrlf = ChrB(13) & ChrB(10) '����һ�����ֽڵĻس����з� 
binSub = DB2SB("--") '����һ�����ֽڵġ�--���ַ��� 

Set sSource = Server.CreateObject("ADODB.stream") 
sSource.Type = 1 'ָ�������������� adTypeBinary=1,adTypeText=2 
sSource.Mode = 3 'ָ����ģʽ adModeRead=1,adModeWrite=2,adModeReadWrite=3 
sSource.open
    dim   content     
    content=request.binaryread(request.totalbytes) 
sSource.Write content '��ȡ���������� 

sSource.Position = 0 
binData = sSource.Read 

posB = InStrB(binData,binSub) '����--��λ�� 
posB = InStrB(posB,binData,vCrlf) + 2 '+2�Ǽ���س����з�����ĳ��� 
posB = InStrB(posB,binData,DB2SB("name=""")) + 6 

Set dControl = Server.CreateObject("Scripting.Dictionary") '������������ؼ�����Ϣ 

Do Until posB=6 'ѭ����������ؼ������� 
posE = InStrB(posB,binData,DB2SB("""")) 
strName = Bin2Str(sSource,posB,posE-posB) '�ؼ����� 

posB = posE + 1 'ָ���ƶ�����"���ĺ��� 
posE = InStrB(posB,binData,vCrlf) 

If InStrB(midb(binData,posB,posE-posB),DB2SB("filename=""")) > 0 Then '�жϳɹ���ʾ����һ��file�� 
posB = InStrB(posB,binData,DB2SB("filename=""")) + 10 
posE = InStrB(posB,binData,DB2SB("""")) 
If posE>posB Then '�ϴ��ļ� 
strFileName = GetFileName(Bin2Str(sSource,posB,posE-posB)) '�ļ��� 
posB = InStrB(posB,binData,DB2SB("Content-Type:")) + 14 
posE = InStrB(posB,binData,vCrlf) 
strContentType = Bin2Str(sSource,posB,posE-posB) '�ļ����� 

posB = posE + 4 '����ط��������У�����ο������ԭʼ���������� 
posE = InStrB(posB,binData,binSub) 

posFB = posB 
posFL = posE-posB-1 
strValue = strFileName & "," & strContentType & "," & posFB & "," & posFL '�����ļ������ļ����͡��ڶ��������ݵ���ʼλ�úͳ��� 
Else 
strValue = "" 
End If 
Else '�������ؼ���ֵ 
posB = posE + 4 '����ط��������У�����ο������ԭʼ���������� 
posE = InStrB(posB,binData,vCrlf) 
strValue = Bin2Str(sSource,posB,posE-posB) 
End If 


dControl.Add strName,strValue '��������������ӵ������� 

posB = posE + 2 

posB = InStrB(posB,binData,vCrlf) + 2 
posB = InStrB(posB,binData,DB2SB("name=""")) + 6 
Loop 

filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&right(strFileName,4)    

bigpicdescribe = dControl.Item("kind") 
arFile = dControl.Item("mefile") 
If arFile <> "" Then '���ļ������ϴ� 
arFile = Split(arFile,",") '��������ļ������ļ����͵����ݷ������ 
RESPONSE.Write(arFile(3))
strFileName = arFile(0) '�ļ��� 
strContentType = arFile(1) '�ļ����� 
posFB = arFile(2) '��ʼλ�� 
posFL = arFile(3) '���� 
sSource.Position = posFB-1
binData = sSource.Read(posFL) '�ӱ������л�ȡͼƬ�Ķ��������� 

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
    'response.write("�ѳɹ��ϴ�,<a href='uploadpic.asp'>�������</a>")      

sSource.Close 
Set sSource = nothing 
end if

%> 
<div style="left:100px; top:50px">���ύ����Ϣ</div>
<div style="left:100px; top:70px">ͼƬ���ࣺ<%=bigpicdescribe%></div>
<div style="left:100px; top:90px">ͼƬ����Ϊ<%=filename%></div>
<div style="left:100px; top:110px">�ѳɹ��ϴ�,<a href='uploadbigpic.asp'>�������</a></div>
</body> 
</HTML> 
