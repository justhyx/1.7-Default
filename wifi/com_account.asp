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
<table width="500" border="0" cellpadding="4" cellspacing="1" style="margin-left:50px" >
  <tr>
    <td colspan="3">&nbsp; </td>
  </tr>
  <tr>
  <td><a href="computer02.asp" target="frmright"><span style="font-weight:bold;">���Թ����¼</span></a></td>
  <td><a href="computer03.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">����ʹ����ϸ</span></a></td>
  <td><a href="computer04.asp" target="frmright"><span style="font-weight:bold;">����ά����¼</span></a></td>
  </tr>
</table>
<%
mac=Trim(request.QueryString("mac"))
ev=Trim(request.QueryString("event"))
if ev<>"broken" then
call createRs(rs,"select * from ����ʹ����ϸ where mac='"& mac &"' and ����ʹ��ʱ�� is null",1,3)
if not rs.eof then
%>
<table width="700" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="4" style="font-weight:bold;">����ʹ������</td>
  </tr>
  <tr>
    <td width="89">ʹ����</td>
	<td width="215"><%=rs("ʹ����")%></td>
    <td width="79">�����ַ</td>
	<td width="243"><%=rs("MAC")%></td>
  </tr>
   <tr>
    <td>IP��ַ</td>
	<td><%=rs("ipdress")%></td>
    <td>��IP</td>
	<td><%=rs("˫IP")%></td>
  </tr>
   <tr>
    <td>�칫��</td>
	<td><%=rs("�칫��")%></td>
    <td>ʹ�ò���</td>
	<td><%=rs("ʹ�ò���")%></td>
  </tr>
   <tr>
    <td>������</td>
	<td><%=rs("������")%></td>
    <td>��ʼʹ��ʱ��</td>
	<td><%=rs("��ʼʹ��ʱ��")%></td>
  </tr>
   <tr>
    <td>����ʹ��ʱ��</td>
	<td><%=rs("����ʹ��ʱ��")%></td>
    <td>��ע</td>
	<td><%=rs("��ע")%></td>
  </tr>
</table>

<%
end if
call closeRS(rs)
end if
call createRs(rs,"select * from ����ʹ����ϸ where mac='"& mac &"' and ����ʹ��ʱ�� is not null",1,3)
if not rs.eof then
%>
<p>&nbsp;</p>
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







