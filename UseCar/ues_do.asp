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
<form id="form1" name="form1" method="post" action="check.asp?id=<%= Trim(Request.QueryString("id")) %>&action=do_use">
  <p>&nbsp;</p>
  <% call CreateRs(rs,"select * from usecar where id="&Trim(Request.QueryString("id")),1,1) %>
  <table width="800" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#337abb">
      <td colspan="6" align="center" style="color:#FFFFFF; font-size:16px; font-weight:bold;">�ó���Ϣ¼��</td>
    </tr>
    
    <tr>
      <td>������</td>
      <td><%= rs("usepeople") %>&nbsp;</td>
      <td>���벿��</td>
      <td><%= rs("cardepartment") %></td>
      <td>��������</td>
      <td><%= rs("reason") %></td>
    </tr>
    <tr>
      <td>������</td>
      <td><label>
        <input name="serverp" type="text" value="<%= rs("serverp") %>" />
      </label></td>
      <td>�복����</td>
      <td><input type="text" name="carstyle" value="<%= rs("carstyle") %>" /></td>
      <td>�˷�</td>
      <td><input type="text" name="do_use" value="<%= rs("do_use") %>" /></td>
    </tr>
    <tr>
      <td>����������</td>
      <td><label>
        <input name="fhbc" type="text" value="<%= rs("fhbc") %>" />
      </label></td>
      <td>���ٹ�·��</td>
      <td><input type="text" name="gslf" value="<%= rs("gslf") %>" /></td>
      <td>���ж��</td>
      <td><label>
        <input name="ddxh" type="text" value="<%= rs("ddxh") %>" />
      </label></td>
    </tr>
    <tr>
      <td>�ϼ��˷�</td>
      <td><input type="text" name="hjyf" value="<%= rs("hjyf") %>" /></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center"><input type="submit" name="Submit" style="background-image:url(../img/button1_n.jpg); font-weight:bold; color:#FFFFFF; height:30px; width:80px;" value="ȷ     ��" /></td>
    </tr>
  </table>
</form>
</body>
</html>
