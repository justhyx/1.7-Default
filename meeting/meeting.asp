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
<form id="form1" name="form1" method="post" action="chkmeeting.asp">  

<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr>
      <td colspan="4" align="center" style="font-weight:bold;">������ʹ�������</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>���벿��</td>
      <td><label>
        <input name="M_Department" type="text" id="M_Department" value="<%= session("Department")%>" />
      </label></td>
      <td>������</td>
      <td><input name="M_usename" type="text" id="M_usename" value="<%= session("UserName")%>" /></td>
    </tr>
    <tr>
      <td>ʹ��ʱ��</td>
      <td><input name="M_usetime" type="text" id="M_usetime" onClick="WdatePicker()" style="hand();" />
        <img onclick="WdatePicker({el:$dp.$('StartTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">&nbsp;
        <label>
        <input name="M_usetime_h" type="text" id="M_usetime_h" size="4" />
        ʱ        </label></td>
      <td>����ʱ��</td>
      <td><input name="M_overtime" type="text" id="M_overtime" onClick="WdatePicker()" />
      <img onclick="WdatePicker({el:$dp.$('OverTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">&nbsp;
      <label>
<input name="M_overtime_h" type="text" id="M_overtime_h" size="4" /> 
ʱ      </label></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>�ͻ�����</td>
      <td><input name="M_guest" type="text" id="M_guest" /></td>
      <td>Ԥ������</td>
      <td><input name="M_number" type="text" id="M_number" /></td>
    </tr>
    <tr>
      <td>�豸����</td>
      <td colspan="3"><label>
        <input name="M_device" type="checkbox" id="M_device" value="ͶӰ��" />
      ͶӰ�� 
      <input name="M_device1" type="checkbox" id="M_device1" value="���Ӱװ�" />
      ���Ӱװ�</label></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>���������</td>
      <td><label>
        <input name="intention" type="text" id="intention" size="8" />
      �Ż�����</label></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center"><label>
        <input type="submit" name="Submit" value="�ᡡ��" />
        <input name="action" type="hidden" id="action" value="add" />
      </label></td>
    </tr>
  </table>
</form>
  <!--#include file="mettuse.asp" -->
</body>
</html>
