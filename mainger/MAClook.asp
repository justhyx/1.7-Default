<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="98">&nbsp;</td>
      <td width="665">&nbsp;</td>
      <td width="31">&nbsp;</td>
      <td width="31">&nbsp;</td>
      <td width="31">&nbsp;</td>
      <td width="31">&nbsp;</td>
      <td width="31">&nbsp;</td>
      <td width="32">&nbsp;</td>
    </tr>
    <tr>
      <td>输入使用人查询：</td>
      <td><label>
        <input name="TxtKey" type="text" id="TxtKey" />
      </label></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<table width="950" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#F1F1F1">
  <tr>
    <td>序号</td>
    <td>使用人</td>
    <td>部门</td>
    <td>职位</td>
    <td>MAC地址</td>
    <td>IP地址</td>
    <td>PPOE地址</td>
    <td>修改时间</td>
  </tr>
   <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=10
'定义总记录数，总页数，当前页数，记数器变量
Dim ThisRsCount,ThisPageCount,ThisCurrentPage,i
'从URL中获取页码Page，并判断有效性
ThisCurrentPage=Request("Page")
If Request("Page")="" Then
	ThisCurrentPage=1
ElseIf IsNumeric(ThisCurrentPage) Then
	ThisCurrentPage=int(ThisCurrentPage)
Else
	ThisCurrentPage=1
End If
If ThisCurrentPage<1 Then ThisCurrentPage=1
%>
<%
'定义查询关键字变量，从URL或表单中获取查询的关键字TxtKey，并组合成模糊查询SQL语句
Dim TxtKey
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from UseMacAddress Order By ID DESC",1,1)
Else
	call CreateRs(rs,"select * from UseMacAddress Where UseName like '%"&TxtKey&"%' Order By ID DESC",1,1)
End If	
%>
<%
'获取总记录数
ThisRsCount=Rs.RecordCount
'设置每页显示多少条记录
Rs.PageSize=ThisPageSize
'获取总页数
ThisPageCount=Rs.PageCount
'判断当前页码的有效性
If ThisCurrentPage>ThisPageCount Then ThisCurrentPage=ThisPageCount
'设置当前页码
Rs.AbsolutePage=ThisCurrentPage
%>
<%
'循环读取数据
i=0
Do While Not Rs.EOF and i<ThisPageSize
%>
  <tr bgcolor="#FFFFFF">
    <td><%= i %>&nbsp;</td>
    <td><%= rs("UseName") %>&nbsp;</td>
    <td><%= rs("UseDepartment") %>&nbsp;</td>
    <td><%= rs("UsePosition") %>&nbsp;</td>
    <td><%= rs("UseMac") %>&nbsp;</td>
    <td><%= rs("UseIp") %>&nbsp;</td>
    <td><%= rs("UsePPPOEIp") %>&nbsp;</td>
    <td><%= rs("AddTime") %>&nbsp;</td>
  </tr>
    <%
	i=i+1
	Rs.MoveNext
Loop	
%>
</table>
<p>&nbsp;</p>
<table width="760" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>总共<%= ThisRsCount %>条记录,<%= ThisPageSize %>条记录/页,总共<%= ThisPageCount %>页,当前第<%= ThisCurrentPage %>页</td>
    <td align="center">
	<% If ThisCurrentPage=1 Then %>
		第一页
	<% Else %>
		<a href="?Page=1&TxtKey=<%=TxtKey%>" class="blue">第一页</a>
	<% End If %>
	</td>
    <td align="center">
	<% If ThisCurrentPage=1 Then %>
		上一页
	<% Else %>
		<a href="?Page=<%=ThisCurrentPage-1%>&TxtKey=<%=TxtKey%>" class="blue">上一页</a>
	<% End If %>
	</td>
    <td align="center">
	<% If ThisCurrentPage=ThisPageCount Then %>
		下一页
	<% Else %>
		<a href="?Page=<%=ThisCurrentPage+1%>&TxtKey=<%=TxtKey%>" class="blue">下一页</a>
	<% End If %>
	</td>
    <td align="center">
	<% If ThisCurrentPage=ThisPageCount Then %>
		最后一页
	<% Else %>
		<a href="?Page=<%=ThisPageCount%>&TxtKey=<%=TxtKey%>" class="blue">最后一页</a>
	<% End If %>
	</td>
  </tr>
</table>
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>
