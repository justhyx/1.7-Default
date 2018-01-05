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
<table width="500" border="0" cellpadding="4" cellspacing="1" style="margin-left:50px" >
  <tr>
    <td colspan="3">&nbsp; </td>
  </tr>
  <tr>
  <td><a href="computer02.asp" target="frmright"><span style="font-weight:bold;">电脑购入记录</span></a></td>
  <td><a href="computer03.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">电脑使用明细</span></a></td>
  <td><a href="computer04.asp" target="frmright"><span style="font-weight:bold;">电脑维护记录</span></a></td>
  </tr>
</table>
<%
mac=Trim(request.QueryString("mac"))
ev=Trim(request.QueryString("event"))
if ev<>"broken" then
call createRs(rs,"select * from 电脑使用明细 where mac='"& mac &"' and 结束使用时间 is null",1,3)
if not rs.eof then
%>
<table width="700" border="1" align="center" cellpadding="8" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="center" colspan="4" style="font-weight:bold;">电脑使用详情</td>
  </tr>
  <tr>
    <td width="89">使用人</td>
	<td width="215"><%=rs("使用人")%></td>
    <td width="79">物理地址</td>
	<td width="243"><%=rs("MAC")%></td>
  </tr>
   <tr>
    <td>IP地址</td>
	<td><%=rs("ipdress")%></td>
    <td>副IP</td>
	<td><%=rs("双IP")%></td>
  </tr>
   <tr>
    <td>办公室</td>
	<td><%=rs("办公室")%></td>
    <td>使用部门</td>
	<td><%=rs("使用部门")%></td>
  </tr>
   <tr>
    <td>责任人</td>
	<td><%=rs("责任人")%></td>
    <td>开始使用时间</td>
	<td><%=rs("开始使用时间")%></td>
  </tr>
   <tr>
    <td>结束使用时间</td>
	<td><%=rs("结束使用时间")%></td>
    <td>备注</td>
	<td><%=rs("备注")%></td>
  </tr>
</table>

<%
end if
call closeRS(rs)
end if
call createRs(rs,"select * from 电脑使用明细 where mac='"& mac &"' and 结束使用时间 is not null",1,3)
if not rs.eof then
%>
<p>&nbsp;</p>
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







