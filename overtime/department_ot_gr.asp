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
</table>
<table width="800" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td>NO</td>
    <td>姓名</td>
    <td>部门</td>
    <td>平时</td>
    <td>周末</td>
    <td>日期</td>
  </tr>
  <%
  month_txt=Trim(Request.QueryString("month_txt"))'year(date())&right("0"&month(date()),2) 
  sql="SELECT CONVERT(char(6), overtime.start_time_1, 112) AS month_txt, overtime.noweek,overtime.isweek,group_txt.group_name, HUDSON_User.UserName, overtime.start_time_1 FROM overtime INNER JOIN HUDSON_User ON overtime.username = HUDSON_User.UserName INNER JOIN  manager_txt ON HUDSON_User.over_l = manager_txt.manager_id INNER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id WHERE (CONVERT(char(6), overtime.start_time_1, 112) = '"& month_txt &"') AND (HUDSON_User.id = '"& Trim(Request.QueryString("user_id")) &"')"
  call CreateRs(rs,sql,1,1)
  k=1
  z=0
  m=0
  do while not rs.eof
   %>
  <tr>
    <td><%= k %>&nbsp;</td>
    <td><%= rs("username") %>&nbsp;</td>
    <td><%= rs("group_name") %>&nbsp;</td>
    <td><%= rs("noweek") %>&nbsp;</td>
    <td><%= rs("isweek") %>&nbsp;</td>
    <td><%= rs("start_time_1") %>&nbsp;</td>
  </tr>
 <%
 z= z + cdbl(rs("noweek"))
 m= m + cdbl(rs("isweek"))
 rs.movenext
 k=k+1
 loop
  %>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right">总计：</td>
    <td><%= z %>&nbsp;</td>
    <td><%= m %>&nbsp;</td>
    <td><%= z + m %>&nbsp;</td>
  </tr>
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
</body>
</html>
