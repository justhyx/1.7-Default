<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<%
dim jhsx(99),bh(99),jhsl(99),i,j,k 
jt=Request.ServerVariables("REMOTE_ADDR")
call createRs(rs,"select * from ��̨�� where IPD_ip='"&jt&"'",1,1)
jtbh=rs("ji")
call closeRs(rs)
call CreateRs(rs,"select * from �ƻ��װ� where equip_name='"&jtbh&"' and okfolg='"&int(1)&"'",1,1)
if rs.eof then
response.write("�˻�̨û�мƻ�")
response.End()
end if
i=0
do while not rs.eof
jhsx(i)=rs("dictate_no")
bh(i)=rs("goods_name")
jhsl(i)=rs("prd_qty")
rs.movenext
i=i+1
loop
call closeRs(rs)
 %>
<body>
<form id="form1" name="form1" method="post" action="chk.asp?event=start">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#000000">
    <tr>
      <td colspan="5" align="center" bgcolor="#FFFFFF">�ƻ�һ</td>
    </tr>
    <tr>
      <td width="13%" bgcolor="#FFFFFF">�ƻ�����</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= date() %>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="4" bgcolor="#FFFFFF">����ƻ�</td>
      <td width="12%" bgcolor="#FFFFFF">����</td>
      <td width="19%" bgcolor="#FFFFFF"><input name="bf" type="text" id="bf" size="16" /></td>
      <td width="9%" bgcolor="#FFFFFF">���</td>
      <td width="47%" bgcolor="#FFFFFF"><input name="bb" type="text" id="bb" size="16" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">�ƻ���</td>
      <td bgcolor="#FFFFFF"><input name="js" type="text" id="js" size="16" /></td>
      <td bgcolor="#FFFFFF">��ҵԱ</td>
      <td bgcolor="#FFFFFF"><input name="zyy" type="text" id="zyy" size="16" /></td>
    </tr>	
    <tr>
      <td bgcolor="#FFFFFF">��̨</td>
      <td bgcolor="#FFFFFF"><input name="jt" type="text" id="jt" value="<%= jtbh %>" size="16" /></td>
      <td bgcolor="#FFFFFF">װ����</td>
      <td bgcolor="#FFFFFF"><input name="zxs" type="text" id="zxs" size="16" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">����</td>
	  <td bgcolor="#FFFFFF"><input name="textfield11" type="text" size="16" /></td>
	  <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><input name="dd" type="text" id="dd" size="16" /></td>
    </tr>
    <% 
		call CreateRs(rs,"select * from �ƻ��װ� where equip_name='"&jtbh&"' and okfolg='"&int(0)&"'",1,1)
		k=0
		do while not rs.eof		
	 %>
	<tr>
      <td bgcolor="#FFFFFF">�ƻ�<%= k+1 %>(�Ѿ����)</td>
      <td bgcolor="#FFFFFF">����<%= jhsx(k)  %></td>
      <td bgcolor="#FFFFFF"><%= bh(k) %></td>
      <td bgcolor="#FFFFFF">����</td>
      <td bgcolor="#FFFFFF"><%= jhsl(k) %></td>
	</tr>
	<% 
	rs.movenext
	k=k+1
	loop
	 %>
  </table>
 <center><br/>
    <input style="height:40px; width:71px;" name="" type="submit" value="    �� ʼ    " />
</center>
</form>
</body>
</html>
<% 
call closeConn()
 %>
