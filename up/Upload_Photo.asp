<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<SCRIPT language=javascript>
<!--#include file="../connet/conn.asp" -->
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("��ѡ��Ҫ�ϴ����ļ�");
		document.form1.FileName.focus();
    	return false;
  	}
}
</SCRIPT>
<% 
if session("UserName")="����" or session("UserName")="�ź�ƻ" then
 %>
</head>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-left: 0px;
}
-->
</style>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css">
<body>
<form action="Upfile_Photo.asp" method="post" name="form1" onSubmit="return check()" enctype="multipart/form-data">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <table width="600" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F0F0F0">
    <tr>
      <td colspan="2" align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">QA/QC��Ʒ����ͼƬ�ϴ�</td>
    </tr>
    <tr>
      <td width="20%" align="right" bgcolor="#FCFCFC">��Ʒ��ţ�</td>
      <td width="80%" bgcolor="#FCFCFC"><label>
     <% 
	 	call CreateRs(rs,"select * from QA_Check order by id desc",1,1) 
	 	QA_ID=rs("QA_ID")
		id=rs("id")
	 %>
	 
	  <input name="QA_ID" type="text" id="QA_ID" value="<%= QA_ID %>">
	  <% call closeRs(rs)%>
      </label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">������</td>
      <td bgcolor="#FCFCFC"><label>
        <input name="QA_name" type="text" id="QA_name">
      </label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">QA/QC�ĵ���</td>
      <td bgcolor="#FCFCFC"><input name="FileName" height="30" type="FILE" size="30"></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">&nbsp;</td>
      <td bgcolor="#FCFCFC"><label></label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">&nbsp;</td>
      <td bgcolor="#FCFCFC"><input type="submit" name="Submit" value="�� ��" ></td>
    </tr>
  </table>
  <table width="600" border="0" align="center" cellpadding="8" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
      <td align="center"><a href="chk.asp?id=<%= id %>&objct=big">ɾ��</a></td>
    </tr>
  </table>
</form>
<% call CreateRs(rs,"select * from QA_image where QA_id='"&QA_ID&"'",1,1) %>
<table width="600" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr bgcolor="#86C4C4" style="font-weight:bold;">
    <td>���</td>
    <td>����</td>
    <td>ͼƬ�鿴</td>
    <td>����</td>
  </tr>
  <% do while not rs.eof %>
  <tr bgcolor="#FFFFFF">
    <td><%= rs("QA_id") %>&nbsp;</td>
    <td><%= rs("QA_name") %></td>
    <td><a href="../<%= rs("QA_idimage") %>">����鿴</a>&nbsp;</td>
    <td><a href="chk.asp?id=<%= rs("id")%>&objct=small">ɾ��</a></td>
  </tr>
  <% 
  rs.movenext
  loop
  call closeRs(rs)
  call closeConn()
   %>
</table>
<% 
	else
	response.Write"<script> alert('ʹ���û���׼ȷ���뷵��'); history.back(); </script>"
	response.End()
end if
 %>
</body>
</thml>

