<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<%
dim mac
mac=request.QueryString("mac")
call createRs(rs,"select* from ���Թ����¼ where MAC='"& mac &"'",1,3)
if not rs.eof and rs("״̬")="����" then
%>
<form id="form01" name="form01" method="post" action="comchk.asp?event=add02">
<table><tr><td>&nbsp;</td></tr></table>
<table width="300" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="2" style="font-weight:bold;">ʹ���˷���</td>
  </tr>
  <tr>
    <td>ʹ�õ���</td>
	<td><input type="text" id="mac" name="mac" value="<%=mac%>"readonly="true" size="32"></td>
  </tr>
  <tr>
    <td>IP����</td>
	<td><input type="text" id="use_ip" name="use_ip" size="32"></td>
  </tr>
  <tr>
    <td>��IP</td>
	<td><input type="text" id="db_ip" name="db_ip" size="32"></td>
  </tr>
   <tr>
    <td>ʹ����</td>
	<td><input type="text" id="use_man" name="use_man" size="32"></td>
  </tr>
   <tr>
    <td>ʹ�ò���</td>
	<td><input type="text" id="use_dep" name="use_dep" size="32"></td>
  </tr>
   <tr>
    <td>������</td>
	<td><input type="text" id="use_boss" name="use_boss" size="32"></td>
  </tr>
     <tr>
    <td>�칫��</td>
	<td><input type="text" id="use_place" name="use_place" size="32"></td>
  </tr>
   <tr>
    <td>��ע</td>
	<td><input type="text" id="note" name="note" size="32"></td>
  </tr>
</table>
<table width="11%"style="margin-left:630px">
<tr><td height="5px">&nbsp;</td></tr>
<tr><td><input type="submit" name="ȷ��" id="ok" value="�� ��"></td></tr></table>
</form>
<%
else
response.Write"<script> alert('δ�ҵ��˵��Ի�˵�������ʹ�ã�������ѡ����ԣ�'); history.back(); </script>"
end if
call closeRs(rs)
call createRs(rs,"select * from ����ʹ����ϸ where mac='"& mac &"' and ����ʹ��ʱ�� is not null",1,3)
if not rs.eof then
%>
<table width="700" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="4" style="font-weight:bold;">��ʷʹ������</td>
  </tr>
  <tr>
  <td>NO</td>
  <td>ʹ����</td>
  <td>IP</td>
  <td>��ʼʹ��ʱ��</td>
  <td>����ʹ��ʱ��</td>
</tr>
<%
i=1
do while not rs.eof
%>
<tr>
<td><%=i%></td>
<td><%=rs("ʹ����")%></td>
<td><%=rs("ipdress")%></td>
<td><%=rs("��ʼʹ��ʱ��")%></td>
<td><%=rs("����ʹ��ʱ��")%></td>
</tr>
<%
i=i+1
rs.movenext
loop
%>
</table>
<%
end if
call closeRs(rs)
call createRs(rs,"select * from ����ά����¼ where mac='"& mac &"'",1,1)
if not rs.eof then
%>
<table width="700" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="5" style="font-weight:bold;">ά����¼</td>
  </tr>
  <tr>
  <td>NO</td>
  <td>ʹ����</td>
  <td>ά����</td>
  <td>ά������</td>
  <td>ά����Ŀ</td>
  <td>ά��ԭ��</td>
</tr>
<%
i=1
do while not rs.eof
%>
<tr>
<td><%=i%></td>
<td><%=rs("comuser")%></td>
<td><%=rs("ά����")%></td>
<td><%=rs("ά������")%></td>
<td><%=rs("ά����Ŀ")%></td>
<td><%=rs("ά��ԭ��")%></td>
</tr>
<%
i=i+1
rs.movenext
loop
%>
</table>
<%
end if
call closeRs(rs)
call closeConn()
%>
</body>
</html>







