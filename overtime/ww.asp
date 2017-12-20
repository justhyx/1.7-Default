<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Server.ScriptTimeOut=900 %>
<!--#include file="../connet/conn.asp" -->
<% 
	call CreateRs(rs1,"select * from hudson_user",1,1)
	do while not rs1.eof
		call CreateRs(rs,"update overtime set user_l='"& rs1("over_l") &"' where username='"& rs1("username") &"'",1,3)
		if rs.eof then
		rs1.movenext
		loop
		call closeRs(rs)
		else
		rs1.movenext
		loop
		call closeRs(rs)
		end if
		call closeRs(rs1)
	call closeConn()
	
 %>
 <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
</body>
</html>
