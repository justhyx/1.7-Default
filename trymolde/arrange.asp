<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>试模联络单管理</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
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
yy=year(now)
mm=getFormatDateString(month(now))
dd=getFormatDateString(int(day(now))+1)
dastr=yy&"-"&mm&"-"&dd
 %>
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
<% if session("UserName")="张云" then  %>
<body>
          <form id="form1" name="form1" method="post" action="arrange.asp">
            <label>
            <input name="TxtKey" type="text" id="TxtKey" class="noprint" />
            </label>
            <label>
            <input type="submit" name="Submit" value="查 询" class="noprint" />
            </label>
            <label class="noprint">输入日期查询 例：2012-06-07 注：输入填单日期 | <a href="arrange_a.asp">当日试模安排</a></label>
          </form>  
  <table id="mytable" width="2000" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center" style="font-weight:bold;">
      <td>试模类型</td>
      <td>型/品番</td>
      <td>生准要望机台</td>
      <td>穴数</td>
      <td>确认机台</td>
      <td>安排时间</td>
      <td>备注</td>
      <td>纳期</td>
      <td>试作数量</td>
      <td>材料</td>
      <td>备注</td>
      <td>日期</td>
      <td>材料编号</td>
      <td>模具担当</td>
      <td>模具状态</td>
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
Dim TxtKey
TxtKey=Request("TxtKey")
If TxtKey="" Then
		call CreateRs(rs,"select * from trymolde where 纳期='"&dastr&"' Order By 生管确认机台 ASC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where 纳期 like '%"&TxtKey&"%' Order By 生管确认机台 ASC",1,1)
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
	<form id="form<%= i %>" name="form<%= i %>" method="post" action="chk.asp?action=arrange">
    <tr>
      <td><%= rs("试模类型") %> | <%= rs("类型") %></td>
      <td><%= rs("品番") %>&nbsp;</td>
      <td><%= rs("生准要望机台") %>&nbsp;</td>
      <td><%= rs("穴数") %>&nbsp;</td>
      <td><input name="okji" type="text" id="okji" value="<%= rs("生管确认机台") %>" size="10" /></td>
      <td><input name="arrangetime" type="text" id="arrangetime" value="<%= rs("安排时间") %>" size="10" /></td>
      <td><input name="sgnote" type="text" id="sgnote" value="<%= rs("生管备注") %>" size="25" />
        <input name="id" type="hidden" id="id" value="<%= rs("id") %>" />
      <input type="submit" name="Submit2" value="提交" /></td>
      <td><%= rs("纳期") %>&nbsp;</td>
      <td><%= rs("试作数量") %></td>
      <td><%= rs("材料") %></td>
      <td><%= rs("备注") %></td>
      <td><%= rs("addtime") %></td>
      <td><label>
        <%= rs("材料编号") %></label></td>
      <td><%= rs("模具担当") %></td>
      <td><label><%= rs("模具状态") %></label></td>
    </tr>
	</form>
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
end if
%>