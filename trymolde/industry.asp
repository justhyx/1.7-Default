<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
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
    <td colspan="4" bgcolor="#909090"><font color="#FFFFFF">������ģ�ƻ�</font></td>
  </tr>
  <tr>
    <td>NO&nbsp;</td>
    <td>����</td>
    <td>����</td>
    <td>��̨</td>
  </tr>
  <% 
	call CreateRs(rs,"select * from trymolde Where ���� ='"&dastr&"' Order By ģ�ߵ��� DESC ",1,1)
		if rs.eof or rs.bof then
			response.write("û������")
			response.End()
		end if
k=1
do while not rs.eof
 %>
  <tr>
    <td><%= k %>&nbsp;</td>
 	<td><%= rs("Ʒ��") %></td>  
    <td><%= rs("��������") %></td>
    <td><%= rs("��׼Ҫ����̨") %></td>  
  </tr>
  <% 
rs.movenext
k=k+1
loop
call closeRs(rs)
 %>
  <tr>
    <td colspan="4" align="center"><a href="#">ǰ����ģ</a> | <a href="#">������ģ</a> | <a href="#" target="_blank">������ģ</a></td>
  </tr>
  <%
Call closeConn()
 %>
</table>
</body>
</html>
