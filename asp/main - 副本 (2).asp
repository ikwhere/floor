<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>地板预览界面</title>
<!--#include file="dbconn.asp"-->
<script>
function change_flag( objectstr ){
	var img = document.getElementById("bigbackground");
    img.style.backgroundImage="url(../Flags/"+objectstr+")"
	}
function change_flagbig( objectstr ){
	var img = document.getElementById("bigpic");
    img.style.backgroundImage="url(../Flags/"+objectstr+")"
	img.style.width="453"
	img.style.height="355"
	}
</script>
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
.title {position:absolute; width:200px; height:115px; z-index:4; top:0px; left:0px;
       }
.background {background-repeat:no-repeat;background-image:url(../pic/123.gif);position:relative; width:468px; height:290px; z-index:0; left:8px; right:-5px
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
<table width="629" height="430" border="0" align="center" cellpadding="0" cellspacing="0" id="maintable">
  <caption>
  <span class="STYLE2">地板预览界面</span>
  </caption>
  <tr>
    <td height="25" colspan="8">
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
        <div  style="background:#000;width:100%;position:absolute; width:80px; height:35px; z-index:2; top:16px ;left:<%=pointx1%>px;;-moz-opacity:0.4; filter:alpha(opacity=30);margin:1px"></div>
		<div style="background-image:url(../pic/hall.gif);text-align:center;line-height:34px;position:absolute;top:15px ;left:<%=pointx%>px; width:80px; height:30px; z-index:2; border:0px solid #555;filter:alpha(opacity=90);"><a href="#" onClick="change_flagbig('<%=rsbigpic("bigpicpath")%>')"><%=rsbigpic("bigpicdiscribe")%></a></div>
         <%
		    pointx=pointx+110
		   	
		  rsbigpic.movenext
		 loop
		 %>
		 <div style="background:#F6E9B9;text-align:LEFT;line-height:34px;position:absolute;top:319px ;left:15px; width:452px; height:50px; z-index:2; border:0px solid #555;filter:alpha(opacity=50);">
		 <span style="color:#0000FF ">
		 <%
		  Flag=1
		 IF request.QueryString("picpath")<>"" THEN
		    Flag=0
		    set rs1=server.CreateObject("adodb.Recordset")
	        strsql="select * from picpath where picpath='"&request.QueryString("picpath")&"'"
			rs1.open strsql,db,1
		 %>
		 描述：<%=rs1("picdiscribe")%>
		 <% END IF
		 %>
		 </span></div>
	  </div></td>
  </tr>
  <tr>
    <td height="15">&nbsp;</td>
    <td colspan="3" rowspan="4"><span <%IF Flag=1 THEN %>style="background-color:#FBF5DF; "<% ELSE  picpath="url(../Flags/"&rs1("picpath")&")" %> style=" background-image:<%=picpath%>" <%END IF%> id="bigbackground" ><img src="" width="453" height="355"  name="bigpic"/></span></td>
    <td height="15">&nbsp;</td>
    <td width="20" rowspan="4"><img src="../pic/picleft.gif" width="20" height="358" /></td>
    <td width="98" rowspan="2" background="../pic/picmid.gif"><div align="center">
        <%
	J=page_no+1
	K=page_no-1
	IF page_no>1 THEN %>
        <a href='main.asp?page_no=<%=K%>'><img src="../pic/arrowup.gif" width="30" height="30" border="0" /></a>
        <%
	END IF%>
    </div></td>
    <td width="28" rowspan="4"> <img src="../pic/picright.gif" width="21" height="358" /></td>
  </tr>
  <tr>
    <td rowspan="2">&nbsp;</td>
    <td width="6" rowspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="279" align="center" valign="middle" background="../pic/picmid.gif">
      <%
			DO WHILE NOT rs.eof  AND I>0
		        I=I-1
   %>
      <a href="main.asp?picpath=<%=rs("picpath")%>#" onClick="change_flag('<%=rs("picpath")%>')"> <img src="../Flags/<%=rs("picpath")%>" width="70" height="50"  border="0"></a>
    <%       response.Write " <BR><BR><BR>"
	         rs.movenext
		   LOOP
	%>    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="35" background="../pic/picmid.gif" style="margin-bottom:-5px ">
      <div align="center" >
        <%IF page_no<rs.pagecount THEN 
	
	%>
        <a href='main.asp?page_no=<%=J%>'><img src="../pic/arrowdown.gif" width="30" height="30" border="0" /></a>
        <%END IF%>
    </div></td>
  </tr>
  <tr>
    <td width="14" height="35" rowspan="2"><div align="left"></div></td>
    <td width="129">&nbsp;</td>
    <td width="154">&nbsp;</td>
    <td width="162" height="19">&nbsp;</td>
    <td height="19" colspan="3">&nbsp; </td>
  </tr>
  <tr>
    <td valign="middle"><img src="../pic/contract.gif" width="30" height="30" /><% IF request.QueryString("picpath")<>"" THEN%><%=rs1("contract")%><%END IF%></td>
    <td width="154" valign="middle"><img src="../pic/tel.gif" width="30" height="30" /><% IF request.QueryString("picpath")<>"" THEN%><%=rs1("tel")%><%END IF%></td>
    <td width="162" height="39" valign="middle"><img src="../pic/mail.gif" width="30" height="30" /><% IF request.QueryString("picpath")<>"" THEN%><%=rs1("email")%><%END IF%></td>
    <td height="39" colspan="3" valign="middle"><img src="../pic/address.gif" width="30" height="30" /><% IF request.QueryString("picpath")<>"" THEN%><%=rs1("addr")%><%END IF%></td>
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

