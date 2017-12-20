<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<%
response.cookies("comsql")=""
response.cookies("comsql").Expires=dateadd("h",24,now())
%>
<body>
<table width="500" border="0" align="left" cellpadding="4" cellspacing="1" bordercolor="#000000" style="margin-left:50px">
  <tr>
    <td colspan="3">&nbsp; </td>
  </tr>
  <tr>
  <td style="font-weight:bold;"><a href="computer02.asp" target="frmright">电脑购入记录</a></td>
  <td style="font-weight:bold;"><a href="computer03.asp" target="frmright">电脑使用明细</a></td>
  <td style="font-weight:bold;"><a href="computer04.asp" target="frmright">电脑维护记录</a></td>
  </tr>
</table>
</body>
</html>







