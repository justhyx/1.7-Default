<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
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
<form id="form1" name="form1" method="post" action="">
  <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="106">输入依赖人查询</td>
      <td width="314"><label>
        <input name="TxtKey" type="text" id="TxtKey" />
        <input type="submit" name="Submit" value="查 询" />
      </label></td>
	  <td align="right">	  	  <input type="button" name="Submit222" onclick="javascript:window.location.href='itevent.asp';" style="font-weight:bold; height:30px; width:120px;" value="设备维护依赖 " /></td>
      <td width="227">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<table border="0" style="margin-left:30px;" cellpadding="4" cellspacing="1" bordercolor="#000000" width="95%" bgcolor="#F1F1F1" id="mytable">
  <tr>
    <td colspan="9" align="center" style=" font-weight:bold;">设备维护依赖查询&nbsp;</td>
  </tr>

  <tr align="center" style="font-weight:bold;">
    <td >编号</td>
    <td >依赖人</td>
    <td >依赖时间</td>
    <td >现象描述</td>
	<td >紧急度</td>
	<td >处理状态</td>	
    <td >审核状态</td>
	<% if session("Power")<999 then %>
	<td >操作</td>
	<% end if %>
	<% if session("user_l")="11" then%>
	<td >依赖明细</td>
	<%
	end if
	%>
  </tr>

   <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=30
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
TxtKey=Trim(Request.Form("TxtKey"))
If TxtKey="" Then
	call CreateRs(rs,"select * from EPevent where 状态='未处理' Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from EPevent Where comuser like '%"&TxtKey&"%' and 状态='未处理' Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.Write("没有此记录")
			response.End()
		end if
End If	
'循环读取数据
i=0
Do While Not rs.eof
call createRs(rsm,"select *from HUDSON_User where UserID='"& rs("comuser") &"'",1,1)
if not rsm.eof then
comuser=rsm("username")
else
comuser=""
end if
call closeRs(rsm)
%>  
  <tr align="center">
    <td ><%= i+1 %>&nbsp;</td>
    <td><%= comuser %>&nbsp;</td>
    <td><%= rs("依赖时间") %>&nbsp;</td>
    <td><%= rs("现象描述") %>&nbsp;</td>
	<td><%= rs("jj")  %>&nbsp;</td>
    <td><%= rs("状态")  %>&nbsp;</td>	
	<% if session("user_l")="19" then%>
	<td><%= rs("chk") %></td>
	<% if session("Power")<999 then %>
	<td><a href="chk_ok.asp?id=<%= rs("id") %>">审核</a></td>
	<% end if %>
	<td><a href="queryIT.asp?id=<%= rs("id")%>">明细查看</a></td>
	<% end if%>
  </tr>
    <%
	i=i+1
	Rs.MoveNext
Loop	
%>
</table>
<p>&nbsp;</p>
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>







