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
<form id="form01" name="form01" method="post" action="comchk.asp?event=change01">
<table><tr><td>&nbsp;</td></tr></table>
<%
id=Trim(request.QueryString("id"))
call createRs(rs,"select * from 电脑购入记录 where id='"& id &"'",1,3)
if not rs.eof then
%>
<table width="300" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="2" style="font-weight:bold;">购入信息更改</td>
  </tr>
  <tr>
    <td>购入日期</td>
	<td><input type="text" id="buy_time" name="buy_time" onClick="WdatePicker()" style="hand();" value="<%=rs("购入日期")%>" size="32"></td>
  </tr>
   <tr>
    <td>购入单价</td>
	<td><input type="text" id="buy_mon" name="buy_mon" value="<%=rs("购入单价")%>" size="32"></td>
  </tr>
   <tr>
    <td>供应商</td>
	<td><input type="text" id="buy_sup" name="buy_sup" value="<%=rs("供应商")%>" size="32"></td>
  </tr>
   <tr>
    <td>物理地址</td>
	<td><input type="text" id="mac" name="mac" value="<%=rs("MAC")%>"size="32"></td>
  </tr>
   <tr>
    <td>类型</td>
	<td><input type="text" id="buy_mold" name="buy_mold" value="<%=rs("类型")%>" size="32"></td>
  </tr>
   <tr>
    <td>位置</td>
	<td><input type="text" id="place" name="place" value="<%=rs("位置")%>"size="32"><input type="hidden" id="buy_id" name="buy_id" value="<%=rs("id")%>"size="32"></td></tr>
	<%if rs("状态")="闲置" then %>
	<tr>
    <td>状态</td>
	<td><select name="zt">
	<option value="0">当前状态</option>
	<option value="1">报废</option>
	</select>
	</td>
  </tr>
  <% end if%>
</table>
<%else
response.Write"<script> alert('未找到可更改的记录，请重新确认！'); history.back(); </script>"
end if
call closeRs(rs)
call closeConn()
%>
<table width="11%"style="margin-left:630px">
<tr><td><input type="submit" name="确定" id="ok" value="提交"></td></tr></table>
</form>
</body>
</html>






