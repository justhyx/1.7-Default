<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="file:///E|/Transaction Management/css/css.css" rel="stylesheet" type="text/css" />
<% 
Function getFormatDateString(dtNumber)
 Dim strDate
 If isNumeric(dtNumber) Then
  If dtNumber<10 Then
   dtNumber = "0" & dtNumber
  End If
  strDate = dtNumber
 Else
  strDate = "" 
 End if
 getFormatDateString = strDate
End Function
dastr=year(now)&"-"&getFormatDateString(month(now))&"-"&getFormatDateString(day(now))
 %>
</head>
<body>
<table width="100%" border="1" cellpadding="1" cellspacing="0" bordercolor="#030303">
  <tr>
    <td colspan="4" bgcolor="#909090"><font color="#FFFFFF">当日试模计划</font></td>
  </tr>
  <tr>
    <td>NO&nbsp;</td>
    <td>部番</td>
    <td>数量</td>
    <td>机台</td>
  </tr>
  <% 
	call CreateRs(rs,"select * from trymolde Where 纳期 ='"&dastr&"' Order By 模具担当 DESC ",1,1)
		if rs.eof or rs.bof then
			response.write("没有试摸")
			response.End()
		end if
k=1
do while not rs.eof
 %>
  <tr>
    <td><%= k %>&nbsp;</td>
 	<td><%= rs("品番") %></td>  
    <td><%= rs("试作数量") %></td>
    <td><%= rs("生准要望机台") %></td>  
  </tr>
  <% 
rs.movenext
k=k+1
loop
call closeRs(rs)
 %>
  <tr>
    <td colspan="4" align="center"><a href="#">前日试模</a> | <a href="#">当日试模</a> | <a href="#" target="_blank">次日试模</a></td>
  </tr>
  <%
Call closeConn()
 %>
</table>
</body>
</html>
