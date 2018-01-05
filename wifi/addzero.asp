<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
 <script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
</head>
<body>
<form id="form09" name="form09" method="post" action="comchk.asp?event=add01">
<table><tr><td>&nbsp;</td></tr></table>
<table width="300" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="2" style="font-weight:bold;">电脑购入信息</td>
  </tr>
  <tr>
    <td>购入日期</td>
	<td><input type="text" id="buy_date" name="buy_date" onClick="WdatePicker()" value="<%=date()%>"size="32" style="hand();"></td>
  </tr>
   <tr>
    <td>购入单价</td>
	<td><input type="text" id="buy_mon" name="buy_mon" size="32"></td>
  </tr>
   <tr>
    <td>供应商</td>
	<td><input type="text" id="buy_sup" name="buy_sup" size="32"></td>
  </tr>
   <tr>
    <td>物理地址</td>
	<td><input type="text" id="mac" name="mac"size="32"></td>
  </tr>
   <tr>
    <td>类型</td>
	<td><input type="text" id="buy_mold" name="buy_mold" size="32"></td>
  </tr>
     <tr>
    <td>位置</td>
	<td><input type="text" id="place" name="place" size="32"></td>
  </tr>
</table>

<table width="11%"style="margin-left:630px">
<tr><td><input type="submit" name="确定" id="ok" value="提交"></td></tr></table>
</form>
</body>
</html>






