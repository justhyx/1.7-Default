<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>日常IT事务管理</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
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
      <td align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">设备维修改善依赖申报</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
    <tr>
      <td width="15%" align="center" bgcolor="#FFFFFF">依赖人厂牌号</td>
      <td width="20%" bgcolor="#FFFFFF"><label>
        <input name="name" type="text" id="name" size="18" value="<%=session("user_id")%>"/>
      </label></td>
      <td width="12%" bgcolor="#FFFFFF" align="center">依赖部门</td>
      <td width="53%"  bgcolor="#FFFFFF"><input name="IT_Tel" type="text" id="IT_Tel" size="50"/></td>
    </tr>
	    <tr>
	      <td align="center" bgcolor="#FFFFFF">设备编号/名称</td>
	      <td bgcolor="#FFFFFF"><input name="equip_no" type="text" id="equip_no" size="18" /></td>
	      <td bgcolor="#FFFFFF" align="center">依赖内容</td>
	      <td  bgcolor="#FFFFFF"><input name="ylnr" type="text" id="ylnr" size="18" /></td>
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
          <td align="center" bgcolor="#FFFFFF">期望纳期</td>
          <td colspan="7" bgcolor="#FFFFFF"><input name="nq" type="text" id="nq" onClick="WdatePicker()" /> 
            紧急度
            <input name="jj" type="radio" value="一般" checked="checked" />
            一般
            <input type="radio" name="jj" value="较急" />
            较急
            <input type="radio" name="jj" value="特急" />
            特急</td>
        </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">依赖详情/现象描述</td>
      <td colspan="7" bgcolor="#FFFFFF"><textarea name="itdetail" cols="64" rows="8" id="itdetail"></textarea></td>
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







