<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>试模联络单管理</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<script language="javascript">
window.onload=function showtable(){
   var tablename=document.getElementById("mytable");
   var li=tablename.getElementsByTagName("tr");
   for (var i=0;i<=li.length;i++){
    if(i%2==0){
     li[i].style.backgroundColor="#F1F1F1";
     li[i].onmouseout=function(){
         this.style.backgroundColor="#F1F1F1";
      }
    }else{
     li[i].style.backgroundColor="#FFFFFF";
     li[i].onmouseout=function(){
         this.style.backgroundColor="#FFFFFF"
      }
    }
      li[i].onmouseover=function(){
      this.style.backgroundColor="#00CCFF";
      }
      
   }
}
</script>
<body>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50" align="center" style="font-size:24px; font-family:'黑体'; font-weight:bold;">试模联络单</td>
  </tr>
  <tr>
    <td height="20" align="right" style="font-weight:bold;"><a href="print.asp">预览</a></td>
  </tr>
</table>
<table id="mytable" width="1050" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center" style="font-weight:bold;">
    <td>试模类型</td>
    <td>型/品番</td>
    <td>望机台</td>
    <td>穴数</td>
    <td>材料编号</td>
    <td>材料</td>
    <td>试作数量</td>
    <td>纳期</td>
    <td>模具担当</td>
    <td>备注</td>
    <td>日期</td>
	<td>操作</td>
  </tr>
   <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=16
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
	call CreateRs(rs,"select * from trymolde where addtime='"&date()&"' and okfolg='"&int(1)&"' Order By id DESC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where 品番 like '%"&TxtKey&"%' Order By ID DESC",1,1)
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
i=0
Do While Not Rs.EOF and i<ThisPageSize
%>
  <tr>
    <td><%= rs("试模类型") %>&nbsp;</td>
    <td><%= rs("品番") %>&nbsp;</td>
    <td><%= rs("生准要望机台") %>&nbsp;</td>
    <td><%= rs("穴数") %>&nbsp;</td>
    <td><%= rs("材料编号") %>&nbsp;</td>
    <td><%= rs("材料") %>&nbsp;</td>
    <td><%= rs("试作数量") %>&nbsp;</td>
    <td><%= rs("纳期") %>&nbsp;</td>
    <td><%= rs("模具担当") %>&nbsp;</td>
    <td><%= rs("备注") %></td>
    <td><%= rs("addtime") %></td>
	<td><% if session("UserName")="庞宇" or session("UserName")="余艳军" then  %><a href="chk.asp?id=<%= rs("id")%>&action=modify&op=isok&ipdress=trymoldelist">承认</a> | <a href="chk.asp?id=<%= rs("id")%>&action=modify&op=isno&ipdress=trymoldelist" onclick="return confirm('确定要取消这条记录吗？')">取消</a><% end if %></td>	
  </tr>
      <%
	i=i+1
	Rs.MoveNext
Loop	
%>
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
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>
