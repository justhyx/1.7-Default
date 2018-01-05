<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
dim conn,str
	set conn=server.CreateObject("adodb.connection")
		str="provider=microsoft.Jet.OLEDB.4.0;data source="& server.MapPath("db/hudsonwwwroot.mdb")
		conn.open str
		
sub CreateRs(rs,sql,n1,n2)
	set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn,n1,n2
end sub

sub closeRs(rs)
	rs.close
	set rs=nothing
end sub

sub closeConn()
	conn.close
	set conn=nothing
end sub
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-color: #E5E2DA;
}
-->
</style>
<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="396" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
    <td colspan="4" align="center" bgcolor="#E5E2DA" style="font-weight:bold;">会议室使用状况</td>
  </tr>
  <tr>
    <td bgcolor="#E5E2DA">申请人</td>
    <td bgcolor="#E5E2DA">使用时间</td>
    <td bgcolor="#E5E2DA">会议室号</td>
    <td bgcolor="#E5E2DA">状态</td>
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
Dim TxtKey,agree,ArrangeTime
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from meeting Order By id DESC",1,1)
Else
	call CreateRs(rs,"select * from meeting Where M_usename like '%"&TxtKey&"%' Order By ID DESC",1,1)
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
  <tr>
    <td bgcolor="#E5E2DA"><%= rs("M_usename") %>&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_usetime") %>&nbsp;&nbsp;<%= rs("M_usetime_h") %>&nbsp;&nbsp;时　至　<%= rs("M_overtime") %>&nbsp;&nbsp;<%= rs("M_overtime_h") %>&nbsp;&nbsp;时&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_meetingroom") %>&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_flog") %>&nbsp;</td>
  </tr>
        <%
	i=i+1
	Rs.MoveNext	
Loop	
%>
</table>
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>
