<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="300" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center">本周wifi密码</td>
  </tr>
  <tr><% call CreateRs(rs,"select * from wifi where wifidate='"& date() &"'",1,1) %>
    <td align="center"><%= rs("password") %>&nbsp;</td>
  </tr><% 
  call closeRs(rs)
  call closeConn()
   %>
</table>
</body>
</html>







