<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
id=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from EPevent where id='"& id &"'",1,3)
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
      <td align="center" bgcolor="#86C4C4"  style="color:#FFFFFF; font-weight:bold;">�豸ά�����������걨</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�����ˣ�</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("comuser") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">��������</td>
      <td bgcolor="#FFFFFF"><%= rs("ά����Ŀ") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�ֻ�/�̺ţ�</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("phone") %><%= rs("tele") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">��������</td>
      <td bgcolor="#FFFFFF"><%= rs("nq")  %>&nbsp;</td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">����ʱ�䣺</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("����ʱ��") %></td>
      <td width="12%" rowspan="2" align="right" bgcolor="#FFFFFF">ά������:</td>
      <td rowspan="2" bgcolor="#FFFFFF"><textarea name="jy" cols="40" rows="4" id="jy"></textarea>
        <input type="hidden" name="whid" id="whid" value="<%=id%>" />
        <input type="hidden" name="ipdress" id="ipdress" value="<%=ipdress%>"></td>
    </tr>

    <tr>
      <td align="right" bgcolor="#FFFFFF">������</td>
      <td bgcolor="#FFFFFF"><%= rs("jj") %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">��������/����������</td>
      <td bgcolor="#FFFFFF"><%= rs("��������") %>&nbsp;</td>
      <td align="right" bgcolor="#FFFFFF">ά����</td>
      <td bgcolor="#FFFFFF"><select name="whr" id="whr">
        <option value="0">��ѡ��</option>
        <option value="2">�ܽ�ɽ</option>
        <option value="3">������</option>
        <option value="4">����ΰ</option>
	<option value="4">���ȫ</option>
	<option value="4">����ƽ</option>
      </select>      </td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>
        <input name="Submit" type="submit" class="tableborder" onclick="return confirm('ȷ�ϴ������')" value="�� �� �� ��" />
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






