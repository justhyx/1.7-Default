<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
if session("Power")<>9998 then
response.end()
end if
dim action,UserID,UserPassword,UserName,UserSex,Department,Position,Phone,Tel,Power,addTime
action=Trim(Request.QueryString("action"))
if action="add" then
txttitle="�����û�"
UserID=""
UserPassword1=""
UserPassword=""
UserName=""
UserSex=""
Department=""
Position=""
Phone=""
Tel=""
Power=""
rsid=""
useraction="adduser"
end if
if action="modify" then
txttitle="�޸��û�����"
call CreateRs(rs,"select * from HUDSON_User where id="&Trim(Request.QueryString("id")),1,1)
UserID=rs("UserID")
UserName=rs("UserName")
UserSex=rs("UserSex")
Department=rs("Department")
Position=rs("Position")
Phone=rs("Phone")
Tel=rs("Tel")
Power=rs("Power")
useraction="modify"
rsid=Trim(Request.QueryString("id"))
call closeRs(rs)
end if
 %>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="checkUser.asp?id=<%= rsid %>&action=<%= useraction %>">
  <p>&nbsp;</p>
  <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr>
      <td colspan="6" align="center" bgcolor="#86C4C4" style="font-weight:bold; color:#FFFFFF;"><%= txttitle %>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">���ƺ�</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="UserID" type="text" id="UserID" value="<%= UserID %>" />
      </label></td>
      <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><input name="UserName" type="text" id="UserName" value="<%= UserName %>" /></td>
      <td bgcolor="#FFFFFF">ְλ</td>
      <td bgcolor="#FFFFFF"><input name="Position" type="text" id="Position" value="<%= Position %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">�ܡ���</td>
      <td bgcolor="#FFFFFF"><input name="UserPassword" type="password" id="UserPassword" /></td>
      <td bgcolor="#FFFFFF">�Ա�</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="UserSex" type="radio" value="��" checked="checked" />
      </label>
      ��
      <input type="radio" name="UserSex" value="Ů" />
      Ů</td>
      <td bgcolor="#FFFFFF">�ֻ���</td>
      <td bgcolor="#FFFFFF"><input name="Phone" type="text" id="Phone" value="<%= Phone %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">ȷ������</td>
      <td bgcolor="#FFFFFF"><input name="UserPassword1" type="password" id="UserPassword1" /></td>
      <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="Department" id="Department">
          <option value="��Ʒ������">��Ʒ������</option>
          <option value="ע�ܼ���">ע�ܼ���</option>
		  <option value="���������">���������</option>
		  <option value="Ӫҵ�ɹ���">Ӫҵ�ɹ���</option>
		  <option value="���������">���������</option>
  		  <option value="�ֿ�����">�ֿ�����</option>
		  <option value="���ο�">���ο�</option>
		  <option value="������">������</option>
		  <option value="������">������</option>
		  <option value="�����">�����</option>
		  <option value="�����">�����</option>
        </select>
      </label></td>
      <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><input name="Tel" type="text" id="Tel" value="<%= Tel %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">�볧ʱ��</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="adddate" type="text" id="adddate" value="<%= date() %>" />
      </label></td>
      <td>���</td>
      <td><select name="user_l" id="user_l">
        <option value="1">��Ʒ������</option>
        <option value="9">ע�ܼ���</option>
        <option value="6">���������</option>
        <option value="8">Ӫҵ�ɹ���</option>
        <option value="4">���������</option>
        <option value="3">������</option>
        <option value="7">������</option>
        <option value="5">�����</option>
        <option value="2">�����</option>
		<option value="20">��</option>
      </select></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">Ȩ ��</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="Power" id="Power">
          <option value="9999">��ͨ�û�</option>
          <option value="999">�೤</option>
		  <option value="99">����</option>
          <option value="9">����</option>
          <option value="7">����</option>
		  <option value="6">�ܾ���</option>
          <option value="5">����</option>
        </select>
      </label></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="ȷ����" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
<% call closeConn() %>