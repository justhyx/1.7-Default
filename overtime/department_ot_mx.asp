<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="1024" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="800" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="4" align="center">平时加班时间</td>
  </tr>
  <tr>
    <td>NO</td>
    <td>姓名</td>
    <td>加班时间</td>
    <td>明细</td>
  </tr>
 <% 
 month_txt=Trim(Request.QueryString("month_txt"))'year(date())&right("0"&month(date()),2)
 sql="SELECT HUDSON_User.id,CONVERT(char(6), overtime.start_time_1, 112) AS month_txt,SUM(overtime.noweek) AS total_time, group_txt.group_name,HUDSON_User.UserName FROM overtime INNER JOIN HUDSON_User ON overtime.username = HUDSON_User.UserName INNER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id INNER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id GROUP BY CONVERT(char(6), overtime.start_time_1, 112), overtime.over_l,group_txt.group_name, group_txt.group_id, HUDSON_User.UserName,HUDSON_User.id,manager_txt.manager_name HAVING (group_txt.group_id = '"& request.QueryString("group_id") &"') AND (CONVERT(char(6), overtime.start_time_1, 112)= '"& month_txt &"') AND (manager_txt.manager_name = '"& session("username") &"')"
 call CreateRs(rs,sql,1,1)
 k=1
 do while not rs.eof
  %>
  <tr>
    <td><%= k %>&nbsp;</td>
    <td><%= rs("UserName") %>&nbsp;</td>
    <td><%= rs("total_time") %>&nbsp;</td>
    <td><a href="department_ot_gr.asp?user_id=<%= rs("id") %>&month_txt=<%= month_txt %>">查看</a></td>
  </tr>
 <% 
 rs.movenext
 k=k+1
 loop
  %>
</table>
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="800" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="4" align="center">周末加班时间</td>
  </tr>
  <tr>
    <td>NO</td>
    <td>姓名</td>
    <td>加班时间</td>
  </tr>
  <% 
 'month_txt=year(date())&right("0"&month(date()),2)
 sql="SELECT HUDSON_User.ID, CONVERT(char(6), overtime.start_time_1, 112) AS month_txt,group_txt.group_name, HUDSON_User.UserName, SUM(overtime.isweek) AS zm_total FROM overtime INNER JOIN HUDSON_User ON overtime.username = HUDSON_User.UserName INNER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id INNER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id GROUP BY CONVERT(char(6), overtime.start_time_1, 112), overtime.over_l,group_txt.group_name, group_txt.group_id, HUDSON_User.UserName,HUDSON_User.ID, manager_txt.manager_name HAVING (group_txt.group_id = '"& request.QueryString("group_id") &"') AND (CONVERT(char(6), overtime.start_time_1, 112)= '"& month_txt &"') AND (manager_txt.manager_name = '"& session("username") &"')"
 call CreateRs(rs,sql,1,1)
 k=1
 do while not rs.eof
  %>
  <tr>
    <td><%= k %>&nbsp;</td>
    <td><%= rs("UserName") %>&nbsp;</td>
    <td><%= rs("zm_total") %>&nbsp;</td>
  </tr>
  <% 
 rs.movenext
 k=k+1
 loop
 call CloseRs(rs)
 call CloseConn()
  %>
</table>
</body>
</html>
