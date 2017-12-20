<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {font-size: 24px}
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
<table width="80%" border="0" cellpadding="4" cellspacing="1" style="margin-left:30px" >
<tr><td colspan="5">&nbsp;</td></tr>
  <tr>
  <td width="13%"><a href="computer02.asp" target="frmright"><span style="font-weight:bold;">电脑购入记录</span></a></td>
  <td width="13%"><a href="computer03.asp" target="frmright"><span style="font-weight:bold;">电脑使用明细</span></a></td>
  <td width="13%"><a href="computer04.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">电脑维护记录</span></a></td>
  <td width="52%">&nbsp;</td>
<td width="9%" style="font-weight:bold;"algin="right"><a href="addtwo.asp" target="frmright">添加记录</a></td>
  </tr>
</table>
<table border="0" style="margin-left:30px;" cellpadding="4" cellspacing="1" bordercolor="#000000" width="95%" bgcolor="#F1F1F1" id="mytable">
<tr><td colspan="8" style="margin-left:50px;">
<form id="form1" name="form1" method="post" action="comchk.asp?event=search03">
      <select name="ziduan" id="ziduan">
	  <option value="MAC">维护电脑</option>
      <option value="维护日期">维护日期</option>
	  <option value="维护人">维护人</option>
    </select>
	&nbsp;<input type="text" id="checkin" name="checkin" size="32">
	&nbsp;<input type="submit" name="submit" value="搜 索">
	</form></td></tr>

<%
comsql=request.cookies("comsql")
if comsql="" then
call createRS(rs,"select * from 电脑维护记录",1,1)
else
call createRS(rs,comsql,1,1)
end if
response.cookies("comsql")=""
response.cookies("comsql").Expires=dateadd("h",24,now())
%>

  <tr align="center" style="font-weight:bold;">
    <td width="4%" bgcolor="#FFFFFF">编号</td>
    <td width="7%" bgcolor="#FFFFFF">依赖人</td>
    <td width="13%" bgcolor="#FFFFFF">依赖时间</td>
    <td width="12%" bgcolor="#FFFFFF">依赖事项</td>
    <td width="38%" bgcolor="#FFFFFF">现象描述</td>
	<td width="8%" bgcolor="#FFFFFF">处理状态</td>
    <td width="18%" bgcolor="#FFFFFF">操作</td>
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
'获取总记录数
ThisRsCount=Rs.RecordCount
'设置每页显示多少条记录
Rs.PageSize=ThisPageSize
'获取总页数
ThisPageCount=Rs.PageCount
'判断当前页码的有效性
If ThisCurrentPage>ThisPageCount Then ThisCurrentPage=ThisPageCount 
'设置当前页码
if ThisCurrentPage>0 then
Rs.AbsolutePage=ThisCurrentPage
end if
%>
<%
if not rs.eof then
i=1+(ThisCurrentPage-1)*ThisPageSize
Do While Not Rs.EOF and i<ThisPageSize+1+(ThisCurrentPage-1)*ThisPageSize
%>
<tr>
<td><%=i%></td>
<td><%=rs("comuser")%></td>
<td><%=rs("依赖时间")%></td>
<td><%=rs("维护项目")%></td>
<td><%=rs("现象描述")%></td>
<td><%=rs("状态")%></td>
<td width="18%"><a href="fixlook.asp?id=<%= rs("id")%>">查 看</a>&nbsp;<a href="changetwo.asp?id=<%=rs("id")%>" target="frmright">更 改</a>&nbsp;<a href="comchk.asp?event=del03&id=<%=rs("id")%>" target="frmright">删 除</a></td>
</tr>
<%
rs.movenext
i=i+1
loop
else 
%>
<tr><td colspan="8">未找到符合要求的记录，请重置搜索条件</td></tr>
<%
end if
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
</body>
</html>







