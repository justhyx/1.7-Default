<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #00CCCC}
-->
</style>
</head>
<%  call CreateRs(rs,"select * from meeting where id="&Trim(Request.QueryString("id")),1,1) %>
<body>
<p>&nbsp;</p>
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr>
    <td colspan="4" align="center" style="font-weight:bold;">������ʹ�������</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>���벿��</td>
    <td><span class="STYLE1"><%= rs("M_Department") %>&nbsp;</span></td>
    <td>������</td>
    <td><span class="STYLE1"><%= rs("M_usename") %></span></td>
  </tr>
  <tr>
    <td>ʹ��ʱ��</td>
    <td><span class="STYLE1"><%= rs("M_usetime") %>&nbsp;&nbsp;<%= rs("M_usetime_h") %>&nbsp;&nbsp;ʱ</span></td>
	<td>����ʱ��</td>
    <td><span class="STYLE1"><%= rs("M_overtime") %>&nbsp;&nbsp;<%= rs("M_overtime_h") %>&nbsp;&nbsp;ʱ</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>�ͻ�����</td>
    <td><span class="STYLE1"><%= rs("M_guest") %></span></td>
    <td>Ԥ������</td>
    <td><span class="STYLE1"><%= rs("M_number") %></span></td>
  </tr>
  <tr>
    <td>�豸����</td>
    <td><span class="STYLE1"><%= rs("M_device") %>&nbsp;&nbsp;<%= rs("M_device1") %></span></td>
    <td>ʹ�û�����</td>
    <td><span class="STYLE1"><%= rs("M_meetingroom") %>�Ż�����</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>������״̬</td>
    <td><span class="STYLE1"><%= rs("M_flog") %>&nbsp;</span></td>
    <td>������</td>
    <td><span class="STYLE1"><%= rs("M_manger") %></span></td>
  </tr>
  <tr>
    <td >�ύʱ��</td>
    <td ><span class="STYLE1"><%= rs("addtime") %> </span></td>
    <td >���ʱ��</td>
    <td ><span class="STYLE1"><%= rs("uptime") %></span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td >���������</td>
    <td ><%=  rs("intention")%>&nbsp;</td>
    <td >&nbsp;</td>
    <td >&nbsp;</td>
  </tr>
</table>
<%
if session("Department")="������" then
%>
<form id="form1" name="form1" method="post" action="chkmeeting.asp">
  <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="11%">�����Ұ���״̬</td>
      <td width="89%"><label>
        <select name="M_flog" id="M_flog">
          <option value="Ԥ�����">Ԥ�����</option>
		  <option value="ʹ����">ʹ����</option>          
		  <option value="ʹ�ý���">ʹ�ý���</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td>�����Һ�</td>
      <td><label>
        <select name="M_meetingroom" id="M_meetingroom">
          <option value="1">1�Ż�����</option>
          <option value="2">2�Ż�����</option>
          <option value="3">3�Ż�����</option>
          <option value="4">4�Ż�����</option>
          <option value="5">5�Ż�����</option>
          <option value="6">6�Ż�����</option>
          <option value="7">7�Ż�����</option>
		  <option value="8">8�Ż�����</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td><label>
        <input type="submit" name="Submit" value="�ᡡ��" />
      </label></td>
      <td><input name="action" type="hidden" id="action" value="update" />
      <input name="id" type="hidden" id="id" value="<%= rs("id") %>" /></td>
    </tr>
  </table>
  <!--#include file="mettuse.asp" -->
</form>
<% 
end if
call closeRs(rs)
call closeConn()
 %>
</body>
</html>







