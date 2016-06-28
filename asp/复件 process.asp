<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dbconn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>处理结果</title>
</head>
<% 
dim   contentlen     
    contentlen=request.totalbytes     
    
    if   contentlen>102400   then     
      response.write   "文件太大，超过100k，不允许上传。请返回"     
    else    
    dim   content     
    content=request.binaryread(request.totalbytes)     
    
    
    Function   getByteString(StringStr)     
      getByteString=""   
      For   i=1   to   Len(StringStr)     
          char=Mid(StringStr,i,1)     
        getByteString=getByteString&chrB(AscB(char))     
      Next     
    End   Function     
    Function   getString(StringBin)     
                  getString   =""     
                  For   intCount   =   1   to   LenB(StringBin)     
                    getString   =   getString   &   chr(AscB(MidB(StringBin,intCount,1)))       
                  Next     
          End   Function     
    
    dim   upbeg,upend,lineone,linetwo,linethree,line1,line2,line3     
    upbeg=1     
    upend=instrb(upbeg,content,getbytestring(chr(10)))     
    lineone=midb(content,upbeg,upend-upbeg)     
    upbeg=upend+1     
    line1=lenb(lineone)     
    upend=instrb(upbeg,content,getbytestring(chr(10)))     
    linetwo=midb(content,upbeg,upend-upbeg)     
    upbeg=upend+1     
    line2=lenb(linetwo)     
    upend=instrb(upbeg,content,getbytestring(chr(13)))     
    linethree=midb(content,upbeg,upend-upbeg)     
    line3=lenb(linethree)     
        
    
    dim   pp,checknametemp,checklen,checkname,filename     
    pp=instrb(1,linetwo,getbytestring(chr(46)))     
    checknametemp=rightb(linetwo,line2-pp+1)     
    checklen=instrb(1,checknametemp,getbytestring(chr(34)))     
    checkname=getstring(leftb(checknametemp,checklen-1))     
    filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&checkname     
        
    
    dim   alllen,upstream,upstreamend,file     
    alllen=line1+line2+line3+6     
    set   upstream=server.createobject("adodb.stream")     
    set   upstreamend=server.createobject("adodb.stream")     
    upstream.type=1     
    upstreamend.type=1     
    upstream.open     
    upstreamend.open     
    upstream.write   content     
    upstream.position=alllen     
    file=upstream.read(clng(contentlen-alllen-line1-5))     
    upstreamend.write   file  
	//response.Write(server.MapPath(("..\Flags\")&filename)) 
    upstreamend.savetofile server.MapPath(("..\Flags\")&filename  )   
    upstream.close     
    upstreamend.close     
    set   upstream=nothing     
    set   upstreamend=nothing     
    
	set dbcmd=server.CreateObject("adodb.Command")
	strsql="Insert into picpath(picpath) values ('"&filename&"')"
	dbcmd.ActiveConnection=db
	dbcmd.CommandText=strsql
	dbcmd.Execute
    response.write("已成功上传,<a href='uploadpic.asp'>点击返回</a>")   
    end   if   

%>
<body>
</body>
</html>
