<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<%
dim mac
mac=request.QueryString("mac")
call createRs(rs,"select* from 电脑购入记录 where MAC='"& mac &"'",1,3)
if not rs.eof and rs("状态")="闲置" then
%>
<form id="form01" name="form01" method="post" action="comchk.asp?event=add02">
<table><tr><td>&nbsp;</td></tr></table>
<table width="300" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="2" style="font-weight:bold;">使用人分配</td>
  </tr>
  <tr>
    <td>使用电脑</td>
	<td><input type="text" id="mac" name="mac" value="<%=mac%>"readonly="true" size="32"></td>
  </tr>
  <tr>
    <td>IP分配</td>
	<td><input type="text" id="use_ip" name="use_ip" size="32"></td>
  </tr>
  <tr>
    <td>副IP</td>
	<td><input type="text" id="db_ip" name="db_ip" size="32"></td>
  </tr>
   <tr>
    <td>使用人</td>
	<td><input type="text" id="use_man" name="use_man" size="32"></td>
  </tr>
   <tr>
    <td>使用部门</td>
	<td><input type="text" id="use_dep" name="use_dep" size="32"></td>
  </tr>
   <tr>
    <td>负责人</td>
	<td><input type="text" id="use_boss" name="use_boss" size="32"></td>
  </tr>
     <tr>
    <td>办公室</td>
	<td><input type="text" id="use_place" name="use_place" size="32"></td>
  </tr>
   <tr>
    <td>备注</td>
	<td><input type="text" id="note" name="note" size="32"></td>
  </tr>
</table>
<table width="11%"style="margin-left:630px">
<tr><td height="5px">&nbsp;</td></tr>
<tr><td><input type="submit" name="确定" id="ok" value="提 交"></td></tr></table>
</form>
<%
else
response.Write"<script> alert('未找到此电脑或此电脑正在使用，请重新选择电脑！'); history.back(); </script>"
end if
call closeRs(rs)
call createRs(rs,"select * from 电脑使用明细 where mac='"& mac &"' and 结束使用时间 is not null",1,3)
if not rs.eof then
%>
<table width="700" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="4" style="font-weight:bold;">历史使用详情</td>
  </tr>
  <tr>
  <td>NO</td>
  <td>使用者</td>
  <td>IP</td>
  <td>开始使用时间</td>
  <td>结束使用时间</td>
</tr>
<%
i=1
do while not rs.eof
%>
<tr>
<td><%=i%></td>
<td><%=rs("使用人")%></td>
<td><%=rs("ipdress")%></td>
<td><%=rs("开始使用时间")%></td>
<td><%=rs("结束使用时间")%></td>
</tr>
<%
i=i+1
rs.movenext
loop
%>
</table>
<%
end if
call closeRs(rs)
call createRs(rs,"select * from 电脑维护记录 where mac='"& mac &"'",1,1)
if not rs.eof then
%>
<table width="700" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="5" style="font-weight:bold;">维护记录</td>
  </tr>
  <tr>
  <td>NO</td>
  <td>使用人</td>
  <td>维护人</td>
  <td>维护日期</td>
  <td>维护项目</td>
  <td>维护原因</td>
</tr>
<%
i=1
do while not rs.eof
%>
<tr>
<td><%=i%></td>
<td><%=rs("comuser")%></td>
<td><%=rs("维护人")%></td>
<td><%=rs("维护日期")%></td>
<td><%=rs("维护项目")%></td>
<td><%=rs("维护原因")%></td>
</tr>
<%
i=i+1
rs.movenext
loop
%>
</table>
<%
end if
call closeRs(rs)
call closeConn()
%>
</body>
</html>







