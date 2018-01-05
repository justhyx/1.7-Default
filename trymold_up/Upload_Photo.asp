<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<SCRIPT language=javascript>
<!--#include file="../connet/conn.asp" -->
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("请选择要上传的文件");
		document.form1.FileName.focus();
    	return false;
  	}
}
</SCRIPT>
</head>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-left: 0px;
}
-->
</style>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css">
<body>
<% 
if session("UserName")="部技" then
 %>
<form action="Upfile_Photo.asp" method="post" name="form1" onSubmit="return check()" enctype="multipart/form-data">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <table width="600" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F0F0F0">
    <tr>
      <td colspan="2" align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">试作确认表上传</td>
    </tr>
    <tr>
      <td width="20%" align="right" bgcolor="#FCFCFC">部番：</td>
      <td width="80%" bgcolor="#FCFCFC"><label>
	  <input name="goods_name" type="text" id="goods_name" >
      </label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">备注：</td>
      <td bgcolor="#FCFCFC"><label>
        <input name="remark" type="text" id="remark">
      </label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">试作确认表：</td>
      <td bgcolor="#FCFCFC"><input name="FileName" height="30" type="FILE" size="30"></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">&nbsp;</td>
      <td bgcolor="#FCFCFC"><label></label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC">&nbsp;</td>
      <td bgcolor="#FCFCFC"><input type="submit" name="Submit" value="上 传" ></td>
    </tr>
  </table>
  <table width="600" border="0" align="center" cellpadding="8" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
      <td align="center"></td>
    </tr>
  </table>
</form>
<% 
end if
 %>
<table width="600" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr bgcolor="#86C4C4" style="font-weight:bold;">
    <td>编号</td>
    <td>部番</td>
    <td>备注</td>
    <td>查看</td>
  </tr>
  <% 
  call creaters(rs,"select * from trymold_up",1,1)
  k=1
  do while not rs.eof
   %>
  <tr bgcolor="#86C4C4" style="font-weight:bold;">
    <td><%= k %>&nbsp;</td>
    <td><%= rs("goods_name") %>&nbsp;</td>
    <td><%= rs("remark") %>&nbsp;</td>
    <td> <a href="pic_list.asp?address=<%= rs("pic_address") %>">查看</a>&nbsp;</td>
  </tr>
  <% 
  rs.movenext
  k=k+1
  loop
   %>
</table>
</body>
</thml>








