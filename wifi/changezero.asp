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
<form id="form01" name="form01" method="post" action="comchk.asp?event=change01">
<table><tr><td>&nbsp;</td></tr></table>
<%
id=Trim(request.QueryString("id"))
call createRs(rs,"select * from ���Թ����¼ where id='"& id &"'",1,3)
if not rs.eof then
%>
<table width="300" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="2" style="font-weight:bold;">������Ϣ����</td>
  </tr>
  <tr>
    <td>��������</td>
	<td><input type="text" id="buy_time" name="buy_time" onClick="WdatePicker()" style="hand();" value="<%=rs("��������")%>" size="32"></td>
  </tr>
   <tr>
    <td>���뵥��</td>
	<td><input type="text" id="buy_mon" name="buy_mon" value="<%=rs("���뵥��")%>" size="32"></td>
  </tr>
   <tr>
    <td>��Ӧ��</td>
	<td><input type="text" id="buy_sup" name="buy_sup" value="<%=rs("��Ӧ��")%>" size="32"></td>
  </tr>
   <tr>
    <td>�����ַ</td>
	<td><input type="text" id="mac" name="mac" value="<%=rs("MAC")%>"size="32"></td>
  </tr>
   <tr>
    <td>����</td>
	<td><input type="text" id="buy_mold" name="buy_mold" value="<%=rs("����")%>" size="32"></td>
  </tr>
   <tr>
    <td>λ��</td>
	<td><input type="text" id="place" name="place" value="<%=rs("λ��")%>"size="32"><input type="hidden" id="buy_id" name="buy_id" value="<%=rs("id")%>"size="32"></td></tr>
	<%if rs("״̬")="����" then %>
	<tr>
    <td>״̬</td>
	<td><select name="zt">
	<option value="0">��ǰ״̬</option>
	<option value="1">����</option>
	</select>
	</td>
  </tr>
  <% end if%>
</table>
<%else
response.Write"<script> alert('δ�ҵ��ɸ��ĵļ�¼��������ȷ�ϣ�'); history.back(); </script>"
end if
call closeRs(rs)
call closeConn()
%>
<table width="11%"style="margin-left:630px">
<tr><td><input type="submit" name="ȷ��" id="ok" value="�ύ"></td></tr></table>
</form>
</body>
</html>






