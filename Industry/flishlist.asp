<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
call CreateRs(rs,"select * from ��ҵ�ձ� where ��̨='"&request.Cookies("scjt")&"' and �ӹ�����='"&request.Cookies("bf")&"' and ��ҵ����='"&date()&"'",1,1)
 %>
</head>

<body>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#000000">
  <tr>
    <td colspan="8" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="4">
      <tr>
        <td height="40" align="center" style="font-size:24px; font-weight:bold;"><%= date() %>�����ƻ�&nbsp;</td>
      </tr>
      <tr>
        <td>������<%= session("bf") %></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">�ƻ���</td>
    <td bgcolor="#FFFFFF"><%= rs("�ƻ�����") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">ʵ�ʲ���</td>
    <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">�ƻ�����</td>
    <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">ʵ������</td>
    <td bgcolor="#FFFFFF"><%= rs("ʵ������") %>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">��ʼʱ��</td>
    <td bgcolor="#FFFFFF"><%= rs("��ʼʱ��") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">����ʱ��</td>
    <td bgcolor="#FFFFFF"><%= rs("��ֹʱ��") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">�ж�ʱ��</td>
    <td bgcolor="#FFFFFF"><%= rs("�ж�ʱ��") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">��Чʱ��</td>
    <td bgcolor="#FFFFFF"><%= rs("�ϼ�ʱ��")-rs("�ж�ʱ��") %>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">��������</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">��������</td>
    <td bgcolor="#FFFFFF"><%= rs("������") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">ȡ��</td>
    <td bgcolor="#FFFFFF"><%= request.Cookies("qs") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">��׼ʱ��</td>
    <td bgcolor="#FFFFFF"><%= rs("��׼ʱ��") %>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">��̨</td>
    <td bgcolor="#FFFFFF"><%= rs("��̨") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">����</td>
    <td bgcolor="#FFFFFF"><%= rs("������") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">��ҵԱ</td>
    <td bgcolor="#FFFFFF"><%= rs("��ҵ��") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">����Ч��</td>
    <td bgcolor="#FFFFFF"><%= rs("����Ч��") %>&nbsp;</td>
  </tr>
  <tr>
    <td height="40" colspan="8" align="center" bgcolor="#FFFFFF" style="font-size:24px; font-weight:bold;"><a href="filshset.asp">ȷ���������</a></td>
  </tr>
</table>
</body>
<% 
call closeRs(rs)
call closeConn()
 %>
</html>
