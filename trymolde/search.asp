<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>试模联络单查询</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body>

<table align="right"><tr><td><a href="addtrymolde.asp?event=add">试模联络单添加</a></td></tr></table>
          <form id="form1" name="form1" method="post" action="search.asp">
		  <p>&nbsp;</p>
           <table style="margin-left:100px"><tr><td>
			 <select name="stylename">
			<option value="模具担当">模具担当</option>\
			<option value="品番">型/品番</option>
			<option value="生准要望机台">生准要望机台</option>
			<option value="试模类型">试模类型</option>
			</select> &nbsp;
            <input name="TxtKey" type="text" id="TxtKey" class="noprint" />
            <input type="submit" name="Submit" value="查 询" class="noprint" />
			</td></tr></table>
          </form>  
  <table id="mytable" width="1000" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center" style="font-weight:bold;">
	  <td>编号</td>
      <td>试模类型</td>
      <td>型/品番</td>
      <td>生准要望机台</td>
      <td>纳期</td>
      <td>日期</td>
      <td>模具担当</td>
	  <td>详细信息</td>
    </tr>
    <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=40
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
Dim TxtKey,stylename
stylename=Trim(request.form("stylename"))
TxtKey=Trim(Request.form("TxtKey"))
If TxtKey="" Then
		call CreateRs(rs,"select * from trymolde",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where "+stylename+" like '%"& TxtKey &"%' Order By addtime",1,1)
		if rs.eof or rs.bof then
			response.Write("没有此记录")
			response.End()
		end if
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
i=1+(ThisCurrentPage-1)*ThisPageSize
Do While Not Rs.EOF and i<ThisPageSize+1+(ThisCurrentPage-1)*ThisPageSize
%>
    <tr align="center">
	  <td><%=i%></td>
      <td><%= rs("试模类型") %> | <%= rs("类型") %></td>
      <td><%= rs("品番") %>&nbsp;</td>
      <td><%= rs("生准要望机台") %>&nbsp;</td>     
      <td><%= rs("纳期") %>&nbsp;</td>
      <td><%= rs("addtime") %></td>
      <td><%= rs("模具担当") %></td>
      <td><a href="addtrymolde.asp?numberid=<%=rs("id")%>&event=change">详细信息</a></td>
    </tr>
    <%
		i=i+1
		Rs.MoveNext
		Loop	
	%>
    <tr>	
      <td colspan="15" align="center">&nbsp;</td>
    </tr>
</table>
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
</body>
</html>
<%
Call CloseRs(rs)
call closeConn()
%>