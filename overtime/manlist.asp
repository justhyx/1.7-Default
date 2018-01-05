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
<form id="form1" name="form1" method="post" action="manlist.asp">
  输入日期：
    <label>
  &nbsp;&nbsp;<input name="TxtKey" type="text" id="TxtKey" />
  &nbsp;<input name="" type="submit" value="搜 索" />
  </label> 
    （例：2012-11-5）
<%= session("over_l") %></form>
<form id="form2" name="form2" method="post" action="chk.asp?action=group_ok">
  <table id="mytable" width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr style="font-weight:bold;">
      <td width="60">NO</td>
      <td>部门</td>
      <td>组别</td>
      <td>姓名</td>
	  <td>申请时间</td>
      <td>开始时间</td>
      <td>结束时间</td>
      <td>事由</td>
      <td>非周末</td>
      <td>周末</td>
    </tr>
    <%
'定义查询关键字变量，从URL或表单中获取查询的关键字TxtKey，并组合成模糊查询SQL语句
Dim TxtKey
if session("Power")<8  then
	TxtKey=Request("TxtKey")
		If TxtKey="" Then
			call CreateRs(rs,"select * from overtime order by gr_flog DESC",1,1)
				if rs.eof or rs.bof then
					response.Write"<font color=red>月没有加班</font>"
					response.End()
				end if
		Else
			call CreateRs(rs,"select * from overtime Where start_time_1='"&TxtKey&"' Order By ID DESC",1,1)
				if rs.eof or rs.bof then
					response.Write"<font color=red>本日没有加班</font>"			
					response.End()
				end if
		End If
else
	TxtKey=Request("TxtKey")
		If TxtKey="" Then
			call CreateRs(rs,"select * from overtime where over_l='"&session("over_l")&"' order by gr_flog DESC",1,1)
				if rs.eof or rs.bof then
					response.Write"<font color=red>月没有加班</font>"
					response.End()
				end if
		Else
			call CreateRs(rs,"select * from overtime Where start_time_1='"&TxtKey&"' and over_l= '"&session("over_l")&"' Order By ID DESC",1,1)
				if rs.eof or rs.bof then
					response.Write"<font color=red>本日没有加班</font>"			
					response.End()
				end if
		End If
end if		
	i=1
	do while not rs.eof
	d1=rs("start_time_1")&" "&rs("start_time_2")
	d2=rs("over_time_1")&" "&rs("over_time_2")
	overtime=datediff("n",d1,d2) 
%>
    <tr>
      <td><%= i %></td>
      <td><%= rs("department") %>&nbsp;</td>
      <td><%= rs("gr_flog") %></td>
      <td><%= rs("username") %>&nbsp;</td>
	  <td><%= rs("apply_time") %>&nbsp;</td>
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
		call CreateRs(rs,"select * from overtime where over_l='"&session("over_l")&"' and start_time_1='"&TxtKey&"'",1,1)
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
	%>
    <tr>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
	  <td align="left">&nbsp;</td>
      <td align="right">小计：</td>
      <td>非周末<%= no_week %></td>
      <td>周末<%= is_week %></td>
      <td>合计：<%= is_week + no_week %></td>
    </tr>
  </table>
</form>
</body>
<% 
call closeRs(rs)
call closeConn()
 %>
</html>
