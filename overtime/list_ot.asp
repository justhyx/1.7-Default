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
	if session("Power")=9998 then
	response.redirect("list_ot_m.asp")
	end if
		if mid(session("Power"),4,1)="1"  then
	response.redirect("list_group.asp")
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
<form id="form1" name="form1" method="post" action="list_ot.asp">
  输入月份：
  <label>
  &nbsp;&nbsp;<input name="TxtKey" type="text" id="TxtKey" />
  &nbsp;<input name="" type="submit" value="搜 索" />
  </label>
  <% if session("Power")<999  then %>  
  <a href="departmen_ot_list.asp">查看明细</a>
  <% end if %>
</form>
<form id="form2" name="form2" method="post" action="chk.asp?action=group_ok">
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1" id="mytable">
  <tr style="font-weight:bold;">
    <% if session("Power")<999  then %>
<td width="60">全选 
      <label>
      <input name="chkall" type="checkbox" id="chkall" onclick="CheckAll(this.form)" />
      </label></td>
	 <% end if %>
    <td>部门</td>
    <td>组别</td>
    <td>姓名</td>
		<% if session("Power")<999 then %>
	<td>申请时间</td>
		<% end if %>
    <td>开始时间</td>
    <td>结束时间</td>
    <td>事由</td>
    <td>非周末</td>
    <td>周末</td>
	<% if session("Power")<999 then %>
    <td>操作</td>
	<% end if %>
  </tr>
<%
'定义查询关键字变量，从URL或表单中获取查询的关键字TxtKey，并组合成模糊查询SQL语句
Dim TxtKey
TxtKey=Request("TxtKey")
if session("Power")<999 and session("Power")>9  then
		If TxtKey="" Then
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id where convert(char(4),start_time_1,112)='"& year(date) &"' and over_l='"&session("over_l")&"' and manager_ok='"&int(1)&"'  order by gr_flog DESC",1,1)
			if rs.eof or rs.bof then
				response.Write"<font color=red>审核完成</font>"
				response.End()
			end if
	Else
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id Where convert(char(4),start_time_1,112)='"& year(date) &"' and mm='"& TxtKey &"' and over_l= '"&session("over_l")&"' and manager_ok='"&int(1)&"' and username<>'"& session("UserName") &"' Order By ID DESC",1,1)
			if rs.eof or rs.bof then
				response.Write"<font color=red>审核完成</font>"			
				response.End()
			end if
	End If	
elseif session("Power")>1000 then
	If TxtKey="" Then
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id where convert(char(4),start_time_1,112)='"& year(date) &"' and username='"&session("username")&"' and manager_ok='"&int(0)&"'and mm='"&month(date)&"'",1,1)
			if rs.eof or rs.bof then
				response.End()
			end if
	Else
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id Where convert(char(4),start_time_1,112)='"& year(date) &"' and mm='"& TxtKey &"' and username= '"&session("username")&"' and manager_ok='"&int(0)&"' Order By ID DESC",1,1)
			if rs.eof or rs.bof then			
				response.End()
			end if
	End If
elseif session("Power")<100 and session("Power")>8 then
		If TxtKey="" Then
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id where convert(char(4),start_time_1,112)='"& year(date) &"' and department='"&session("Department")&"' and manager_ok='"&int(1)&"' order by gr_flog DESC",1,1)
			if rs.eof or rs.bof then
				response.Write"<font color=red>审核完成</font>"
				response.End()
			end if
	Else
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id Where convert(char(4),start_time_1,112)='"& year(date) &"' and mm='"& TxtKey &"' and department= '"&session("Department")&"' and manager_ok='"&int(1)&"' Order By ID DESC",1,1)
			if rs.eof or rs.bof then
				response.Write"<font color=red>审核完成</font>"			
				response.End()
			end if
	End If
elseif session("Power")<8 then
		If TxtKey="" Then
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id overtime where convert(char(4),start_time_1,112)='"& year(date) &"'  and manager_ok='"&int(1)&"' order by gr_flog DESC",1,1)
			if rs.eof or rs.bof then
				response.Write"<font color=red>审核完成</font>"
				response.End()
			end if
	Else
		call CreateRs(rs,"SELECT group_txt.group_name AS gr_txt, * FROM overtime INNER JOIN  group_txt ON overtime.user_l = group_txt.group_id Where convert(char(4),start_time_1,112)='"& year(date) &"' and mm='"& TxtKey &"' and manager_ok='"&int(1)&"' Order By ID DESC",1,1)
			if rs.eof or rs.bof then
				response.Write"<font color=red>审核完成</font>"			
				response.End()
			end if
	End If		
end if	
	i=1
	do while not rs.eof
	d1=rs("start_time_2")
	d2=rs("over_time_2")
	overtime=datediff("n",d1,d2) 
	'apy=rs("apply_time")
	'apply_time=year(apy)&"-"&month(apy)&"-"&day(apy)
%> 
  <tr>
    <% if session("Power")<999 then %>
<td><label>
      <input name="cbox<%= i %>" type="checkbox" id="cbox1" value="<%= rs("id")%>" />
    </label></td>
	<% end if %>
    <td><%= rs("department") %>&nbsp;</td>
	<td><%= rs("gr_txt") %></td>
	<td><%= rs("username") %>&nbsp;</td>
			<% if session("Power")<999 then %>
	<td><%= rs("apply_time") %>&nbsp;</td>
		<% end if %>
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
    <td style="font-weight:bold;"><% if session("Power")<999 and rs("manager_ok")=1 then %>
      <a href="chk.asp?id=<%= rs("id")%>&amp;action=manager_ok">承 认</a> | <a href="chk.asp?id=<%= rs("id")%>&amp;action=manager_NG">取 消</a>
    <% end if %></td>
  </tr>
  <%
		i=i+1
		Rs.MoveNext		
		Loop
		call closeRs(rs)
			
	%>
  <tr> 
  <% if session("Power")<999  then %>   
	<td align="left">
<label>
      <input name="rs_i" type="hidden" id="rs_i" value="<%= i %>" />
      <input type="submit" name="Submit" value="批量承认" />
    </label></td>
	<% end if %>	
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="left">&nbsp;</td>
    <td align="right"></td>
    <td></td>
    <td></td>
    <td>合计：</td>
  </tr>
</table>
</form>
</body>
<% 
call closeConn()
 %>
</html>







