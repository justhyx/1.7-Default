<!--#include file="../connet/conn.asp" -->
<% 
dim Department,Power,i,j
i=0
Department = session("Department")
Power = session("Power")
'�ó��ȴ���������
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
<TITLE>��ӭʹ�úͳ��ǳ��������ϵͳ</TITLE>
<link rel="stylesheet" href="Images/Admin.css">
<BODY>

<table width="95%" border="0" cellspacing="1" cellpadding="4">
  <tr>
    <td width="12%">&nbsp;</td>
    <td width="88%">&nbsp;</td>
  </tr>
  <tr>
    <td>������</td>
    <td><%= session("UserName") %>&nbsp;</td>
  </tr>
  <tr>
    <td>���ţ�</td>
    <td><%= session("Department_txt") %>&nbsp;</td>
  </tr>
  <% if Power<999 then  %>
  <tr>
    <td>�ó����������</td>
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







