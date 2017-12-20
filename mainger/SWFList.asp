<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>
<%
call CreateRs(rs,"select * from QA_Check where id="&Trim(Request.QueryString("id")),1,1) 
link=rs("QA_address")
 %>
<body>
<table width="750" border="0" align="center" cellpadding="4" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="750" height="450">
      <param name="movie" value="../<%= link %>" />
      <param name="quality" value="high" />
      <embed src="../<%= link %>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="750" height="450"></embed>
    </object></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<% 
call closeRs(rs)
call closeConn()
 %>
