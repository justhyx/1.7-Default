<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table id="mytable" width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr style="font-weight:bold;">
    <td>&nbsp;</td>
    <td>����</td>
    <td>����</td>
	<td>����ʱ��</td>
	<td>���ʱ��</td>
    <td>��ʼʱ��</td>
    <td>����ʱ��</td>
    <td>����</td>
    <td>����ĩ</td>
    <td>��ĩ</td>
  </tr>
     <%
	call CreateRs(rs,"select * from overtime where convert(char(4),start_time_1,112)='"&year(now)&"' and username='"&Trim(Request.QueryString("pasname"))&"' and mm='"& session("mmlist") &"' and manager_ok='"&int(0)&"' order by ID ASC",1,1)
	i=0
	is_week=0
	no_week=0
	do while not rs.eof
	d1=rs("start_time_2")
	d2=rs("over_time_2")
	overtime=datediff("n",d1,d2) 
%> 
  <tr>
    <td>&nbsp;</td>
    <td><%= rs("department") %>&nbsp;</td>
	<td><%= rs("username") %>&nbsp;</td>
	<td><%= rs("apply_time")%>&nbsp;</td>
	<td><%= rs("verify_time") %>&nbsp;</td>
    <td><%= rs("start_time_1")%>&nbsp;<%= rs("start_time_2")%></td>
    <td><%= rs("over_time_1") %>&nbsp;<%= rs("over_time_2") %></td>
    <td><%= rs("note") %></td>
    <td><%=rs("noweek")  %></td>
    <td><%=rs("isweek")  %></td>
    <td style="font-weight:bold;">&nbsp;</td>
  </tr>
  <%
		i=i+1
		is_week=rs("isweek")+is_week
  		no_week=rs("noweek")+no_week
		Rs.MoveNext		
		Loop	
	%>
  <tr>
    <td colspan="7" align="right">С�ƣ�</td>
    <td><%= no_week %></td>
    <td><%= is_week %></td>
    <td>�ϼƣ�<%= is_week + no_week %></td>
  </tr>
</table>
</body>
<% 
call closers(rs)
call closeConn()
 %>
</html>







