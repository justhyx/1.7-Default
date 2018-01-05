<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<% 
wifidate=date()
for i=1 to 365
randomize
wifipsw=int(99999999*rnd)+100
call CreateRs(rs,"select * from wifi",1,3)
rs.addnew
rs("password")=wifipsw
rs("wifidate")=wifidate+i
rs.update
call closeRs(rs)
next
call closeConn()
 %>
</head>

<body>
</body>
</html>







