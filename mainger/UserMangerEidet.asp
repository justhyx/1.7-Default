<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
if session("power")=9998 or session("username")="���" then
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
Power="9999"
rsid=""
add_date=date()
Position_txt="��ͨ�û�"
useraction="adduser"
manager_txt="��ѡ��"
group_txt="��ѡ��"
department="��ѡ��"
end if
if action="modify" then
txttitle="�޸��û�����"
call CreateRs(rs,"SELECT * FROM HUDSON_User INNER JOIN  group_txt ON HUDSON_User.user_l = group_txt.group_id INNER JOIN  manager_txt ON HUDSON_User.over_l = manager_txt.manager_id  where HUDSON_User.id="&Trim(Request.QueryString("id")),1,1)
UserID=rs("UserID")
UserName=rs("UserName")
UserSex=rs("UserSex")
Department=rs("Department")
Position=rs("Position")
Phone=rs("Phone")
Tel=rs("Tel")
Power=rs("Power")
user_l=rs("user_l")
over_l=rs("over_l")
UserPassword1=rs("UserPassword")
UserPassword=rs("UserPassword")
useraction="modify"
manager_txt=rs("manager_name")
group_txt=rs("group_name")
manager_id=rs("manager_id")
group_id=rs("group_id")
add_date=rs("addTime")
Position_txt=rs("Position")
power=rs("power")
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
      <td bgcolor="#FFFFFF">�ֻ���</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="Phone" type="text" id="Phone" value="<%= Phone %>" />
      </label></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">�ܡ���</td>
      <td bgcolor="#FFFFFF"><input name="UserPassword" value="<%= UserPassword %>" type="password" id="UserPassword" /></td>
      <td bgcolor="#FFFFFF">�Ա�</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="UserSex" type="radio" value="��" checked="checked" />
      </label>
      ��
      <input type="radio" name="UserSex" value="Ů" />
      Ů      </td>
      <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><input name="Tel" type="text" id="Tel" value="<%= Tel %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">ȷ������</td>
      <td bgcolor="#FFFFFF"><input name="UserPassword1" value="<%=UserPassword1%>"type="password" id="UserPassword1" /></td>
      <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="Department" id="Department">
 <%call createRS(rs,"select * from department where is_del='N'",1,1)
	 do while not rs.eof
	  %>		
          <option value="<%=rs("superior")%>"><%=rs("superior")%></option>
       <%
	   rs.movenext
	   loop
	   call closeRs(rs)
	   %>
          <option value="<%= Department %>" selected="selected"><%= Department %></option>
        </select>
      </label></td>
      <td bgcolor="#FFFFFF">��λ</td>
      <td bgcolor="#FFFFFF"><select name="user_l" id="user_l">
        <%
		call createRS(rs,"select * from group_txt where is_del='N'",1,1)
	    do while not rs.eof
	  %>
        <option value="<%=rs("group_id")%>"><%=rs("group_name")%></option>
        <%rs.movenext
	   loop
       call closeRs(rs)
	   %>
        <option value="<%= group_id %>" selected="selected"><%= group_txt %></option>
      </select>
</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">�볧ʱ��</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="adddate" type="text" id="adddate" value="<%= add_date %>" />
      </label></td>
      <td>��������</td>
      <td><select name="over_l" id="over_l" >
	  <%
	  call createRS(rs,"select * from manager_txt t where is_del='N'",1,1)
	  do while not rs.eof
	  %>	  
        <option value="<%=rs("manager_id")%>"><%= rs("manager_name") %></option>
		<%
		rs.movenext
	   loop
	 call closeRs(rs)
	 %>		
        <option value="<%= manager_id %>" selected="selected"><%= manager_txt %></option>
        </select></td>
      <td>���</td>
      <td><label>
      <select name="txt_group" id="txt_group">
      </select>
      </label></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">Ȩ ��</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="Power" id="Power">
          <option value="<%= Power %>" selected="selected"><%= Position_txt %></option>
		  <option value="999">�೤</option>
          <option value="9998">����</option>
		  <option value="99">����</option>
          <option value="9">����</option>
          <option value="7">������</option>
		  <option value="6">�ܾ���</option>
          <option value="5">���³�</option>
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
<% 
else
response.End()
end if
 %>
</html>
<% call closeConn() %>






