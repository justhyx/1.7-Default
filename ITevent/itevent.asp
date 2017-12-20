<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>日常IT事务管理</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>

<body>
<p>&nbsp;</p>
<form id="form1" name="form1" method="post" action="checkIT.asp?action=add">  
  <table width="95%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">日常IT事务依赖申报</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
    <tr>
      <td width="15%" align="center" bgcolor="#FFFFFF">依赖人</td>
      <td width="20%" bgcolor="#FFFFFF"><label>
        <input name="name" type="text" id="name" size="18" value="<%= session("UserName") %>"/>
      </label></td>
      <td width="12%" bgcolor="#FFFFFF" align="center">依赖事项</td>
      <td width="53%"  bgcolor="#FFFFFF"><input name="IT_Tel" type="text" id="IT_Tel" size="50"/>
        <span class="STYLE1">(不超过20字)</span></td>
    </tr>
	    <tr>
      <td width="15%" align="center" bgcolor="#FFFFFF">依赖人手机/短号</td>
      <td width="20%" bgcolor="#FFFFFF"><label>
        <input name="phone" type="text" id="phone" size="18"/>
      </label></td>
      <td width="12%" bgcolor="#FFFFFF" align="center">内线电话</td>
      <td width="53%"  bgcolor="#FFFFFF"><input name="tele" type="text" id="tele" size="18"/>
        <span class="STYLE1">(手机和内线电话必须填写一项)</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">依赖原因</td>
      <td colspan="7" bgcolor="#FFFFFF"><label>
        <textarea name="itevent" cols="64" rows="6" id="itevent"></textarea>
      </label></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">依赖详情/现象描述</td>
      <td colspan="7" bgcolor="#FFFFFF"><textarea name="itdetail" cols="64" rows="6" id="itdetail"></textarea>
	  <span class="STYLE1">(必须填写，不超过120字)</span>
	  </td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>
        <input name="Submit" type="submit" class="tableborder" value="提 交 依 赖" />
      </label></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
<% call closeConn() %>







