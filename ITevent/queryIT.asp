<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
id=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from ����ά����¼ where id='"& id &"'",1,3)
ipdress=request.ServerVariables("REMOTE_ADDR")

 %>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<p>&nbsp;</p>
<form id="form1" name="form1" method="post" action="checkIT.asp?action=IT_query">  
  <table width="95%" border="1" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td align="center" bgcolor="#86C4C4"  style="color:#FFFFFF; font-weight:bold;">�ճ�IT���������걨</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�����ˣ�</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("comuser") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">��������:</td>
      <td colspan="2" bgcolor="#FFFFFF"><%= rs("ά����Ŀ") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�ֻ�/�̺ţ�</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("phone") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">���ߵ绰:</td>
      <td colspan="2" bgcolor="#FFFFFF"><%= rs("tele") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">����ʱ�䣺</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("����ʱ��") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">IT���:</td>
      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="itmind" id="itmind" size="50">
	  <input type="hidden" name="whid" id="whid" value="<%=id%>">
	  <input type="hidden" name="ipdress" id="ipdress" value="<%=ipdress%>"></td>
    </tr>

    <tr>
      <td align="right" bgcolor="#FFFFFF">����ԭ��</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= rs("ά��ԭ��") %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">��������/����������</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= rs("��������") %>&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>
        <input name="Submit" type="submit" class="tableborder" value="ȷ �� �� ��" />
      </label></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
<% 
call closeRs(rs)
call closeConn()
 %>






