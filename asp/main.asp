<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>地板预览界面</title>
<!--#include file="dbconn.asp"-->

<script>
function tabbk_lload(){
	var img = document.getElementById("maintable");
    img.style.backgroundImage="url(../pic/background.jpg)"
	}
</script>
<%
	set rs=server.CreateObject("adodb.Recordset")
	strsql="select * from picpath"
	rs.open strsql,db,1
	If Not rs.bof and Not rs.Eof Then
	    DIM page_no
		IF request.QueryString("page_no")="" THEN
		     page_no=1
	    ELSE
		     page_no=Cint(request.QueryString("page_no"))
		END IF
		
		rs.pagesize=2
		rs.absolutepage=page_no
		DIM I
		I=rs.pagesize
%>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 24px;
	font-weight: bold;
}
.style3 {font-size: 9px}

.STYLE5 {font-size: 14px}
#Layer2 {
	position:absolute;
	width:25px;
	height:384px;
	z-index:2;
	left: 545px;
	top: 115px;
}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>

<body onLoad="tabbk_lload()">

<p>&nbsp;</p>

<%
         Flag=1
		 IF request.QueryString("picpath")<>"" THEN
		    Flag=0
		    set rs1=server.CreateObject("adodb.Recordset")
	        strsql="select * from picpath where picpath='"&request.QueryString("picpath")&"'"
			rs1.open strsql,db,1
			SESSION("picpath")=rs1("picpath")
			SESSION("picdiscribe")=rs1("picdiscribe")
			SESSION("contract")=rs1("contract")
			SESSION("tel")=rs1("tel")
			SESSION("addr")=rs1("addr")
			SESSION("email")=rs1("email")
           END IF
		   
		  Flagbig=1
	      IF request.QueryString("bigpic")<>"" THEN
	         Flagbig=0
			 SESSION("bigpicpath")= request.QueryString("bigpic")
		  END IF
		 
%>
<table width="629" height="446" border="0" align="center" cellpadding="0" cellspacing="0" id="maintable">
  <caption></caption>
  <div id="Layer1" style="position:relative; width:0px; height:0px; z-index:1;">
	  		<%dim pointx,J,pointx1
	     set rsbigpic=server.CreateObject("adodb.Recordset")
	     strsql="select bigpicpath,bigpicdiscribe from bigpic"
		 pointx=33
	     rsbigpic.open strsql,db,1
		 J=0
		do while not (rsbigpic.Eof )
             pointx1=pointx+2	
		%>
        <div  style="background:#000;width:100%;position:absolute; width:80px; height:35px; z-index:2; top:51px ;left:<%=pointx1%>px;;-moz-opacity:0.4; filter:alpha(opacity=30);margin:1px"></div>
		<div style="background-image:url(../pic/hall.gif);text-align:center;line-height:34px;position:absolute;top:50px ;left:<%=pointx%>px; width:80px; height:30px; z-index:2; border:0px solid #555;filter:alpha(opacity=90);">
		<a href="main.asp?bigpic=<%=rsbigpic("bigpicpath")%>#" ><%=rsbigpic("bigpicdiscribe")%></a></div>
         <%
		    pointx=pointx+110
		   	
		  rsbigpic.movenext
		 loop
		 %>
		 <div style="background:#F6E9B9;text-align:LEFT;line-height:34px;position:absolute;top:365px ;left:14px; width:458px; height:50px; z-index:2; border:0px solid #555;filter:alpha(opacity=70);">
		 <span style="color:#0000FF ">
		 <%
		 IF Flag=0 OR SESSION("picpath")<>"" THEN
		    
		
		 %>
		 描述：<%=SESSION("picdiscribe")%>
		 <% END IF
		 %>
		 </span></div>
  </div>
  
  <tr>
    <td height="50" colspan="8"></td>
  </tr>
  <tr>
    <td width="14" height="396" rowspan="5"></td>
    <td height="361" colspan="3" rowspan="3" >
	<%IF Flag=0 OR SESSION("picpath")<>"" THEN 
	picpath="url(../Flags/"&SESSION("picpath")&")" 
	response.Write "<span  style='background-image:"&picpath&"; border-width:thin; width:361; height:3'>"
	 ELSE  
	response.Write "<span  style='background-color:#FBF5DF; border-width:thin;height:3'>"
	END IF
	IF Flagbig=0 OR SESSION("bigpicpath")<>"" THEN
       response.Write "<img src='../Flags/"&SESSION("bigpicpath")&"'border-width:thick; width='457' height='361' name='bigpic'/>"
	 ELSE
	  response.Write " <img src='' width='457' height='361' name='bigpic'/>"
	END IF
	%>
	</span></td>
    <td width="12" height="361" rowspan="3"></td>
    <td width="20" height="361" rowspan="3"><img src="../pic/picleft.gif" width="20" height="361" /></td>
    <td width="98" height="33" background="../pic/picmid.gif"><div align="center">
        <%
	J=page_no+1
	K=page_no-1
	IF page_no>1 THEN %>
        <a href='main.asp?page_no=<%=K%>'><img src="../pic/arrowup.gif" width="30" height="30" border="0" /></a>
        <%
	END IF%>
    </div></td>
    <td width="28" rowspan="3"> <img src="../pic/picright.gif" width="21" height="361" /></td>
  </tr>
  
  <tr>
    <td height="295" align="center" valign="middle" background="../pic/picmid.gif">
      <%
			DO WHILE NOT rs.eof  AND I>0
		        I=I-1
   %>
      <a href="main.asp?picpath=<%=rs("picpath")%>#"> <img src="../Flags/<%=rs("picpath")%>" width="70" height="50"  border="0"></a>
      <%       response.Write " <BR><BR><BR>"
	         rs.movenext
		   LOOP
	%>    </td>
  </tr>
  <tr>
    <td height="33" background="../pic/picmid.gif" style="margin-bottom:-5px ">
      <div align="center" >
        <%IF page_no<rs.pagecount THEN %>
	    <a href='main.asp?page_no=<%=J%>'><img src="../pic/arrowdown.gif" width="30" height="30" border="0" /></a>
        <%END IF%>
    </div></td>
  </tr>
  <tr>
    <td height="15" colspan="6">&nbsp; </td>
  </tr>
  <tr>
    <td width="0" height="20" valign="middle"><span class="STYLE5"><img src="../pic/contract.gif" width="30" height="30" />
      <% IF Flag=0 OR SESSION("contract")<>"" THEN%>
      <%=SESSION("contract")%>
      <%END IF%>
    </span></td>
    <td width="0" height="20" valign="middle"><span class="STYLE5"><img src="../pic/tel.gif" width="30" height="30" />
      <% IF Flag=0 OR SESSION("tel")<>"" THEN%>
      <%=SESSION("tel")%>
      <%END IF%>
    </span></td>
    <td width="0" height="20" valign="middle"><span class="STYLE5"><img src="../pic/mail.gif" width="30" height="30" />
      <% IF Flag=0 OR SESSION("email")<>"" THEN%>
      <%=SESSION("email")%>
      <%END IF%>
    </span></td>
    <td height="20" colspan="3" valign="middle"><span class="STYLE5"><img src="../pic/address.gif" width="30" height="30" />
      <% IF Flag=0 OR SESSION("addr")<>"" THEN%>
      <%=SESSION("addr")%>
      <%END IF%>
    </span></td>
  </tr>
</table>
<p><span class="style3">
  <% 
	   FOR I=1 TO rs.pagecount 
	             IF I=page_no THEN
				      response.Write I & "&nbsp;"
	             ELSE
				      response.Write "<a href='main.asp?page_no="&I&"'>"&I&"</a>&nbsp;"
			     END IF
				 IF I MOD 8=0 THEN
				      response.Write "<BR>"
				 END IF
		NEXT
	END IF
%>
</span></p>
</body>
</html>

