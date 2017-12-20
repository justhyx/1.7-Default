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
    <td><form id="form1" name="form1" method="post" action="">
      请输入查询月份
      <label>
      <input name="txtkey" type="text" id="txtkey" />
      </label>
        <label>
        <input type="submit" name="Submit" value=" 搜 索 " />
        </label>
    格式：例2016年4月 “201604”
    </form>
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="600" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="4" align="center">平时加班工时查看</td>
  </tr>
  <tr>
    <td>NO</td>
    <td>组</td>
    <td>工时</td>
    <td>明细</td>
  </tr>
 <% 
 month_txt=Trim(Request.Form("TxtKey"))
 if month_txt="" then
 month_txt=year(date())&right("0"&month(date()),2)
 end if
 sql="SELECT CONVERT(char(6), overtime.start_time_1, 112) AS month_txt, overtime.over_l,SUM(overtime.noweek) AS total_time, manager_txt.manager_name,group_txt.group_name,group_txt.group_id FROM overtime INNER JOIN HUDSON_User ON overtime.username = HUDSON_User.UserName INNER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id INNER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id GROUP BY CONVERT(char(6), overtime.start_time_1, 112), overtime.over_l, manager_txt.manager_name, group_txt.group_name,group_txt.group_id HAVING (CONVERT(char(6), overtime.start_time_1, 112) = '"&month_txt&"') AND (manager_txt.manager_name = '"& session("username") &"')"
 call CreateRs(rs,sql,1,1)
 k=1
 do while not rs.eof
  %>  
  <tr>
    <td><%= k %>&nbsp;</td>
    <td><%= rs("group_name") %>&nbsp;</td>
    <td><%= rs("total_time") %>&nbsp;</td>
    <td><a href="department_ot_mx.asp?group_id=<%= rs("group_id") %>&month_txt=<%= month_txt %>">查看</a></td>
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
<table width="600" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="4" align="center">周末加班工时查看</td>
  </tr>
  <tr>
    <td>NO</td>
    <td>组</td>
    <td>工时</td>
  </tr>
  <% 
' month_txt=year(date())&right("0"&month(date()),2)
 sql="SELECT CONVERT(char(6), overtime.start_time_1, 112) AS month_txt, overtime.over_l,manager_txt.manager_name, group_txt.group_name, group_txt.group_id,SUM(overtime.isweek) AS zm_total FROM overtime INNER JOIN HUDSON_User ON overtime.username = HUDSON_User.UserName INNER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id INNER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id GROUP BY CONVERT(char(6), overtime.start_time_1, 112), overtime.over_l,manager_txt.manager_name, group_txt.group_name, group_txt.group_id HAVING (CONVERT(char(6), overtime.start_time_1, 112) = '"&month_txt&"') AND (manager_txt.manager_name = '"& session("username") &"')"
 call CreateRs(rs,sql,1,1)
 k=1
 do while not rs.eof
  %>
  <tr>
    <td><%= k %>&nbsp;</td>
    <td><%= rs("group_name") %>&nbsp;</td>
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
