<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="UserManger.asp">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;
        <input name="TxtKey" type="text" id="TxtKey" />
        <input type="submit" name="Submit" value="�� ѯ" />
        �������ֲ�ѯ
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F0F0F0">
  <tr>
    <td colspan="8" align="center" bgcolor="#86C4C4" style="font-weight:bold; color:#FFFFFF;">�û�����</td>
  </tr>
  <tr align="center" style="font-weight:bold;">
  <td bgcolor="#FFFFFF">NO</td>
    <td bgcolor="#FFFFFF">����</td>
    <td bgcolor="#FFFFFF">�û���</td>
    <td bgcolor="#FFFFFF">����</td>
    <td bgcolor="#FFFFFF">��λ</td>
    <td bgcolor="#FFFFFF">����</td>
    <td bgcolor="#FFFFFF">���ʱ��</td>
    <td bgcolor="#FFFFFF">����</td>
  </tr>
   <%  
'	if session("Power")<>9998 then
'	response.end()
'	end if
	Dim TxtKey
	TxtKey=Request("TxtKey")
	If TxtKey="" Then
		call CreateRs(rs,"SELECT * FROM HUDSON_User INNER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id INNER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id Order By HUDSON_User.department",1,1)		
	Else
		call CreateRs(rs,"SELECT * FROM HUDSON_User LEFT OUTER JOIN group_txt ON HUDSON_User.user_l = group_txt.group_id LEFT OUTER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id Where UserName like '%"&TxtKey&"%' Order By HUDSON_User.username",1,1)		
	End If	
	i=1
	Do While Not Rs.EOF 
	%>  
  <tr>
  <td align="center" bgcolor="#FFFFFF"><%= i %>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%= rs("UserName") %>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%= rs("UserID") %>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%= rs("Department") %>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%= rs("group_name") %>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%= rs("manager_name") %>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%= rs("addTime") %></td>
    <td align="center" bgcolor="#FFFFFF"><a href="UserMangerEidet.asp?id=<%= rs("id") %>&action=modify">�޸�</a> | <a href="checkUser.asp?id=<%= rs("id")%>&amp;action=del" onclick="return confirm('ȷ��ɾ����')">ɾ��</a></td>
  </tr>
<%	
	Rs.MoveNext
	i=i+1
	Loop	
%>
</table>
<p>&nbsp;</p>
</body>
</html>
<% 
call closeRs(rs)
call closeConn() %>






