<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
 <script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
<head>

<% 
id=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from ����ά����¼ where id='"& id &"'",1,3)

 %>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form name="form01" id="form01" method="post" action="comchk.asp?event=add03">
  <table width="95%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td align="center"  style="font-weight:bold;">ά����¼���</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">�����ˣ�</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="fm" id="fm"></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">��������:</td>
      <td width="35%"  bgcolor="#FFFFFF"><input type="text" name="thing" id="thing"></td>
    </tr>
	    <tr>
	 <td align="right" bgcolor="#FFFFFF">����ԭ��</td>
      <td  bgcolor="#FFFFFF"><input type="text" name="reason" id="reason" size="50"></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">����ʱ�䣺</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="sd" id="sd" onClick="WdatePicker()" style="hand();"></td>

    </tr>
	    <tr>
		      <td width="15%" align="right" bgcolor="#FFFFFF">ά������:</td>
      <td  bgcolor="#FFFFFF"><input type="text" name="fc" id="fc"></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">ά��ʱ�䣺</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="ft" id="ft" onClick="WdatePicker()" style="hand();"></td>

    </tr>
	    <tr>
			  <td width="15%" align="right" bgcolor="#FFFFFF">״̬:</td>
              <td bgcolor="#FFFFFF">	  <select name="zt">
	  <option value="<%=rs("״̬")%>" selected="selected"><%=rs("״̬")%></option>
	  <option value="δ����">δ����</option>
	  <option value="�Ѵ���" selected="selected">�Ѵ���</option>
	  </select>
	  </td>
              <td width="15%" align="right" bgcolor="#FFFFFF">ά���ˣ�</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="it" id="it"></td>

    </tr>

    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">IT���:</td>
      <td bgcolor="#FFFFFF"><input type="text" name="im" id="im" size="50"></td>
	   <td align="right" bgcolor="#FFFFFF">��ע��</td>
      <td bgcolor="#FFFFFF"><input type="text" name="bz" id="bz"></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">��������/����������</td>
      <td colspan="3" bgcolor="#FFFFFF"><input type="text" name="xx" id="xx" size="100">&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>	  <input type="submit" name="Submit220"  value=" ��  �� " /></label></td>
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




