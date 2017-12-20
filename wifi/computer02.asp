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
<table width="80%" border="0" cellpadding="4" cellspacing="1" style="margin-left:50px" >
<tr><td colspan="5">&nbsp;</td></tr>
  <tr>
  <td width="12%"><a href="computer02.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">电脑购入记录</span></a></td>
  <td width="12%"><a href="computer03.asp" target="frmright"><span style="font-weight:bold;">电脑使用明细</span></a></td>
  <td width="12%"><a href="computer04.asp" target="frmright"><span style="font-weight:bold;">电脑维护记录</span></a></td>
  <td width="55%">&nbsp;</td>
<td style="font-weight:bold;"algin="right"><a href="addzero.asp" target="frmright">添加记录</a></td></tr>
</table>

<table border="0" style="margin-left:50px;" cellpadding="4" cellspacing="1" bordercolor="#000000" width="100%" id="mytable" bgcolor="#F1F1F1">
<tr><td colspan="7" style="margin-left:50px;">
<form id="form1" name="form1" method="post" action="comchk.asp?event=search01">
      <select name="ziduan" id="ziduan">
      <option value="MAC">物理地址</option>
      <option value="购入日期">购入日期</option>
	  <option value="供应商">供应商</option>
      <option value="类型">类型</option>
	  <option value="状态" selected="selected">状态</option>
	  <option value="位置">位置</option>
    </select>
	&nbsp;<input type="text" id="checkin" name="checkin" size="32">
	&nbsp;<input type="submit" name="submit" value="搜 索">
	</form></td></tr>
<td width="6%">编号</td>
<td width="17%">物理地址</td>
<td width="15%">购入日期</td>
<td width="10%">购入单价</td>
<td width="13%">供应商</td>
<td width="10%">类型</td>
<td width="14%">位置</td>
<td>状态</td>
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
	dim comsql
comsql=request.cookies("comsql")
if comsql="" then
call createRS(rs,"select * from 电脑购入记录",1,1)     '这里的两个参数如果是1，3，则Rs.recordCount是不可用的，会返回为-1
else
call createRS(rs,comsql,1,1)
end if
response.cookies("comsql")=""
response.cookies("comsql").Expires=dateadd("h",24,now())
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
<td><%=rs("MAC")%></td>
<td><%=rs("购入日期")%></td>
<td><%=rs("购入单价")%></td>
<td><%=rs("供应商")%></td>
<td><%=rs("类型")%></td>
<td><%=rs("位置")%></td>
<td width="15%">
<% if rs("状态")="闲置" then
%>
<a href="addone.asp?mac=<%=rs("MAC")%>" target="frmright">
<span style="color:#FF0000">
<%=rs("状态")%>
</span>
</a>
<%elseif rs("状态")="使用中" then
%>
<a href="com_account.asp?mac=<%=rs("MAC")%>" target="frmright">
<span style="color:#0000FF">
<%=rs("状态")%>
</span>
</a>
<%
else
%>
<a href="com_account.asp?event=broken&mac=<%=rs("MAC")%>" target="frmright"><%=rs("状态")%></a>
<%
end if
%>
</a>&nbsp;&nbsp;<a href="changezero.asp?id=<%=rs("id")%>" target="frmright">更 改</a>&nbsp;
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
<%
call closeRs(rs)
call closeConn()
%>


</body>
</html>







