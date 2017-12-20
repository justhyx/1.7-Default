<!--#include file="../connet/conn.asp" -->
<% 
dim Department,Power,i,j
i=0
Department = session("Department")
Power = session("Power")
'用车等待事项数量
if Power<9 then
call CreateRs(rs,"select * from UseCar where UrgencyAgree= '"&int(2)&"'",1,1)
do while not rs.eof
i=i+1
rs.movenext
loop
elseif Power<999 then
call CreateRs(rs,"select * from UseCar where CarDepartment='"&Department&"' and Agree= '"&int(2)&"'",1,1)
do while not rs.eof
i=i+1
rs.movenext
loop
call closeRs(rs)
end if
call closeConn()
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<TITLE>欢迎使用和诚锴诚事务管理系统</TITLE>
<link rel="stylesheet" href="Images/Admin.css">
<BODY>

<table width="95%" border="0" cellspacing="1" cellpadding="4">
  <tr>
    <td width="12%">&nbsp;</td>
    <td width="88%">&nbsp;</td>
  </tr>
  <tr>
    <td>姓名：</td>
    <td><%= session("UserName") %>&nbsp;</td>
  </tr>
  <tr>
    <td>部门：</td>
    <td><%= session("Department_txt") %>&nbsp;</td>
  </tr>
  <% if Power<999 then  %>
  <tr>
    <td>用车待处理事项：</td>
    <td><%= i %>&nbsp;</td>
  </tr>
  <tr>
  <% end if %>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</BODY>
</HTML>







