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
<table width="80%" border="0" cellpadding="4" cellspacing="1" style="margin-left:30px" >
<tr><td colspan="5">&nbsp;</td></tr>
  <tr>
  <td width="13%"><a href="computer02.asp" target="frmright"><span style="font-weight:bold;">���Թ����¼</span></a></td>
  <td width="13%"><a href="computer03.asp" target="frmright"><span style="font-weight:bold;">����ʹ����ϸ</span></a></td>
  <td width="13%"><a href="computer04.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">����ά����¼</span></a></td>
  <td width="52%">&nbsp;</td>
<td width="9%" style="font-weight:bold;"algin="right"></td>
  </tr>
</table>
  <table width="95%" border="0" align="center" cellpadding="10" cellspacing="0" bordercolor="#000000">
    <tr>
      <td align="center"  style="font-weight:bold;">�ճ�IT���������걨</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="0"  bordercolor="#000000">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�����ˣ�</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("comuser") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">��������:</td>
      <td width="35%"  bgcolor="#FFFFFF"><%= rs("ά����Ŀ") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�ֻ�/�̺ţ�</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("phone") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">���ߵ绰:</td>
      <td width="35%"  bgcolor="#FFFFFF"><%= rs("tele") %></td>
    </tr>
	    <tr>
	 <td align="right" bgcolor="#FFFFFF">����ԭ��</td>
      <td  bgcolor="#FFFFFF"><%= rs("ά��ԭ��") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">����ʱ�䣺</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("����ʱ��") %></td>

    </tr>
	    <tr>
		      <td width="15%" align="right" bgcolor="#FFFFFF">ά������:</td>
      <td  bgcolor="#FFFFFF"><%= rs("MAC") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">ά��ʱ�䣺</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("ά������") %></td>

    </tr>
	    <tr>
			  <td width="15%" align="right" bgcolor="#FFFFFF">״̬:</td>
      <td bgcolor="#FFFFFF"><%= rs("״̬") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">ά���ˣ�</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("ά����") %></td>

    </tr>

    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">IT���:</td>
      <td bgcolor="#FFFFFF"><%= rs("IT���") %></td>
	   <td align="right" bgcolor="#FFFFFF">��ע��</td>
      <td bgcolor="#FFFFFF"><%= rs("��ע") %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">��������/����������</td>
      <td colspan="3" bgcolor="#FFFFFF"><%= rs("��������") %>&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>	  <input type="button" name="Submit220" onclick="javascript:history.back();" value=" ��  �� " /></label>
	  <input type="button" name="Submit222" onclick="javascript:window.location.href='changetwo.asp?id=+ <%= rs("id") %> +';" value=" ��  �� " />
	    <input type="button" name="Submit222" onclick="javascript:window.location.href='comchk.asp?event=del03&id=+ <%= rs("id") %> +';" value=" ɾ  �� " />
	  </td>
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




