<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
if request.Form("action")="save" then
	call CreateRs(rs,"select * from UseMacAddress",1,3)
		rs.addnew
		rs("UseName")=Trim(Request.Form("UseName"))
		rs("UseDepartment")=Trim(Request.Form("UseDepartment"))
		rs("UsePosition")=Trim(Request.Form("UsePosition"))
		rs("UseMac")=Trim(Request.Form("UseMac"))
		rs("UseIp")=Trim(Request.Form("UseIp"))
		rs("UsePPPOEIp")=Trim(Request.Form("UsePPPOEIp"))
		rs("AddTime")=now()
		rs.update
	call closeRs(rs)
	response.Write"<script> alert('��ӳɹ�'); history.go(-2);</script>"
	response.End()	
end if
 %>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <table width="600" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center">
      <td colspan="4">���ŵ��Թ滮</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">ʹ���ˣ�</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="UseName" type="text" id="UseName" />
      </label></td>
      <td bgcolor="#FFFFFF">���ţ�</td>
      <td bgcolor="#FFFFFF"><input name="UseDepartment" type="text" id="UseDepartment" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">ְ��</td>
      <td bgcolor="#FFFFFF"><input name="UsePosition" type="text" id="UsePosition" /></td>
      <td bgcolor="#FFFFFF">IP��ַ��</td>
      <td bgcolor="#FFFFFF"><input name="UseIp" type="text" id="UseIp" value="<%= Request.ServerVariables("REMOTE_ADDR") %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF"><label></label>
        <label></label></td>
    </tr>
  </table>
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr align="center">
      <td><input type="submit" name="Submit" value="�� ��" />
      <input name="action" type="hidden" id="action" value="save" /></td>
    </tr>
  </table>
</form>
</body>
</html>
