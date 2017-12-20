<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
session("Url")="../Usecar/UseList.asp"
if session("UserName")="" then
	response.Redirect("../login.asp")
end if
 %>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE2 {color: #FF0000}
-->
</style>
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
</head>

<body>
<a href="personnel_list.asp">总务入口</a>
<table id="mytable" width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center" style="font-weight:bold;">
    <td>申请部门</td>
    <td>申请人</td>
    <td><a href="uselist_addtime.asp">申请时间</a></td>
    <td><a href="UseList.asp?op=StartTime">用车时间</a></td>
    <td>目的地</td>
    <td>审核结果</td>
    <td>审核时间</td>
    <td>确认出车</td>
    <td>出车时间</td>
    <td>查看</td>
  </tr>
   <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=21
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
	call CreateRs(rs,"select * from UseCar where delfolg<>'"&int(2)&"' and over_l='"& session("over_l") &"' Order BY ID DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from UseMacAddress Where UseName like '%"&TxtKey&"%'",1,1)
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
  <tr align="center" bgcolor="#FFFFFF">
    <td><%= rs("CarDepartment") %></td>
    <td><%= rs("UsePeople") %>&nbsp;</td>
    <td><%= rs("addtime") %></td>
    <td><%= rs("StartTime") %></td>
    <td><%= rs("Destination")%></td>
    <% if rs("Agree")="0" or rs("UrgencyAgree")="0" then %>
	<td>审核未通过</td>	
	<% elseif rs("Agree")="1" or rs("UrgencyAgree")="1" then %>
	<td>同意</td>	
	<% elseif rs("Urgencyfloy")>0 then  %>
	<td><span class="STYLE2">等待审核</span></td>
	<% end if %>
	<td><%= rs("ArrangeTime") %>&nbsp;</td>
	<% if rs("arrange")="1" then %>
	<td>
	<% if session("UserName")="钱影梅" then %>
	<a href="UseLook.asp?action=arrange&id=<%= rs("id") %>"><font color="#FF0000">派车等待</font></a>
	<% else %>
	<font color="#FF0000">派车等待</font>
	<% end if %>
	</td>
	<% else %>
	<td>安排完成</td>
	<% end if %>
	<td><%= rs("addarrangetime") %>&nbsp;</td>
	<td><a href="UseLook.asp?id=<%= rs("id")%>">查看</a></td>
    <% if rs("Urgency")<>"" then %>
	<td><span class="STYLE2">急&nbsp;</span></td>
	<% end if %>
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
