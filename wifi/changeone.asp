<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
 <script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
</head>

<body>
<form id="form01" name="form01" method="post" action="comchk.asp?event=change02">
<table><tr><td>&nbsp;</td></tr></table>
<%
id=Trim(request.QueryString("id"))
call createRs(rs,"select * from ����ʹ����ϸ where id='"& id &"'",1,3)
if not rs.eof then
call createRs(rsa,"select *from ���Թ����¼ where MAC='"& rs("MAC") &"'",1,1)
%>
<table width="350" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="2" style="font-weight:bold;">ʹ����Ϣ����</td>
  </tr>
  <tr>
    <td>ʹ����</td>
	<td><input type="text" id="use_man" name="use_man" value="<%=rs("ʹ����")%>"size="32" readonly="true"></td>
  </tr>
   <tr>
    <td>�����ַ</td>
	<td><input type="text" id="mac" name="mac" value="<%=rs("MAC")%>" size="32"></td>
  </tr>
   <tr>
    <td>IP��ַ</td>
	<td><input type="text" id="ipdress" name="ipdress" value="<%=rs("ipdress")%>" size="32"></td>
  </tr>
   <tr>
    <td>��IP</td>
	<td><input type="text" id="db_ip" name="db_ip" value="<%=rs("˫IP")%>" size="32"></td>
  </tr>
   <tr>
    <td>�칫��</td>
	<td><input type="text" id="use_place" name="use_place" value="<%=rs("�칫��")%>" size="32"></td>
  </tr>
   <tr>
    <td>ʹ�ò���</td>
	<td><input type="text" id="use_dep" name="use_dep" value="<%=rs("ʹ�ò���")%>" size="32"></td>
  </tr>
   <tr>
    <td>������</td>
	<td><input type="text" id="use_boss" name="use_boss" value="<%=rs("������")%>"size="32"></td>
  </tr>
   <tr>
    <td>��ʼʹ��ʱ��</td>
	<td><input type="text" id="use_timeone" name="use_timeone" onClick="WdatePicker()" style="hand();" value="<%=rs("��ʼʹ��ʱ��")%>"size="32"></td>
  </tr>
   <tr>
    <td>����ʹ��ʱ��</td>
	<td><input type="text" id="use_timetwo" name="use_timetwo" onClick="WdatePicker()" style="hand();"  value="<%=rs("����ʹ��ʱ��")%>" size="32"><input type="hidden" id="buy_id" name="buy_id" value="<%=rsa("id")%>"size="32"><input type="hidden" id="use_id" name="use_id" value="<%=rs("id")%>"size="32"></td>
  </tr>
     <tr>
    <td>��ע</td>
	<td><input type="text" id="note" name="note"  value="<%=rs("��ע")%>"size="32"></td>
  </tr>
</table>
<table width="11%"style="margin-left:630px">
<tr><td height="5px">&nbsp;</td></tr>
<tr><td><input type="submit" name="ȷ��" id="ok" value="�� ��"></td></tr></table>
<%call closeRS(rsa)
end if
call closeRS(rs)
call closeConn()
%>
</form>
</body>
</html>






