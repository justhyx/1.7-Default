<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
	if session("Power")="" then
	response.Redirect("../index.asp")
	end if
 %>
<script language="javascript">
function CheckAll(form2)  {
  for (var i=0;i<form2.elements.length;i++){
    var e = form2.elements[i];
    if (e.name != 'chkall')
	e.checked = form2.chkall.checked; 
   }
  }
</script>
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
<form id="form1" name="form1" method="post" action="list_group.asp">
  输入月份：
  <label>
  &nbsp;&nbsp;<input name="TxtKey" type="text" id="TxtKey" />
  &nbsp;<input name="" type="submit" value="搜 索" />
  </label>
</form>
<form id="form2" name="form2" method="post" action="chk.asp?action=group_ok">
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1" id="mytable">
  <tr style="font-weight:bold;">
    <td>部门</td>
    <td>组别</td>
    <td>姓名</td>
    <td>开始时间</td>
    <td>结束时间</td>
    <td>事由</td>
    <td>非周末</td>
    <td>周末</td>
  </tr>
<%
'定义查询关键字变量，从URL或表单中获取查询的关键字TxtKey，并组合成模糊查询SQL语句
Dim TxtKey

		call CreateRs(rsa,"select username from HUDSON_User where user_l='"&session("user_l")&"' and power>99 order by Power",1,1)
			if rsa.eof or rsa.bof then
				response.End()
			end if
			do while not rsa.eof			
    TxtKey=Request("TxtKey")
	If TxtKey="" Then
	sql="select * from overtime where convert(char(4),start_time_1,112)='"& year(date) &"' and username='"& rsa("username") &"' and manager_ok='"&int(0)&"'and mm='"&month(date)&"' order by cast(start_time_1 as datetime) desc"
	else
	sql="select * from overtime Where convert(char(4),start_time_1,112)='"& year(date) &"' and mm='"& TxtKey &"' and username= '"&rsa("username")&"' and manager_ok='"&int(0)&"' Order By start_time_1"
	end if
		call CreateRs(rs,sql,1,1)
			if not rs.eof then
	i=1
	do while not rs.eof
	d1=rs("start_time_2")
	d2=rs("over_time_2")
	'overtime=datediff("n",d1,d2) 
%> 
  <tr>
    <td><%= rs("department") %>&nbsp;</td>
	<td><%= rs("gr_flog") %></td>
	<td><%= rs("username") %>&nbsp;</td>
    <td><%= rs("start_time_1")%>&nbsp;<%= rs("start_time_2")%></td>
    <td><%= rs("over_time_1") %>&nbsp;<%= rs("over_time_2") %></td>
    <% if rs("other_flog")=1 then
			bcolor="FF0000"
		else
			bcolor="000000"
		end if	
	 %>
	<td style="color:#<%= bcolor %>"><%= rs("note") %></td>
    <td><%=rs("noweek")  %></td>
    <td><%=rs("isweek")  %></td>
  </tr>
  <%
		i=i+1
		Rs.MoveNext		
		Loop
		call closeRs(rs)

		call CreateRs(rs,"select * from overtime where convert(char(4),start_time_1,112)='"& year(date) &"' and username='"&rsa("username")&"' and mm='"&month(date)&"' and manager_ok='"&int(0)&"'",1,1)
			if rs.eof or rs.bof then
			is_week=0
			no_week=0
			end if
			is_week=0
			no_week=0
			do while not rs.eof
			is_week=rs("isweek")+is_week
			no_week=rs("noweek")+no_week
			rs.movenext
			loop
			call closeRS(rs)					
	%>
  <tr> 	
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="right">小计：</td>
    <td><%= no_week %></td>
    <td><%= is_week %></td>
    <td>合计：<%= is_week + no_week %></td>
  </tr>
  <tr><td colspan="8">&nbsp;</td></tr>
  <%
  else
%>
<tr><td colspan="8" align="right"><%=rsa("username")%>本月内没有加班记录！</td></tr>
<tr><td colspan="8">&nbsp;</td></tr>
<%
end if
    rsa.movenext
  loop
%>
</table>
</form>
</body>
<% 
call closeRs(rsa)
call closeConn()
 %>
</html>







