<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
call CreateRs(rs,"select * from 作业日报 where 机台='"&request.Cookies("scjt")&"' and 加工部番='"&request.Cookies("bf")&"' and 作业日期='"&date()&"'",1,1)
 %>
</head>

<body>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#000000">
  <tr>
    <td colspan="8" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="4">
      <tr>
        <td height="40" align="center" style="font-size:24px; font-weight:bold;"><%= date() %>生产计划&nbsp;</td>
      </tr>
      <tr>
        <td>部番：<%= session("bf") %></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">计划数</td>
    <td bgcolor="#FFFFFF"><%= rs("计划数量") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">实际产量</td>
    <td bgcolor="#FFFFFF"><%= rs("产量") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">计划周期</td>
    <td bgcolor="#FFFFFF"><%= rs("周期") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">实际周期</td>
    <td bgcolor="#FFFFFF"><%= rs("实际周期") %>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">开始时间</td>
    <td bgcolor="#FFFFFF"><%= rs("起始时间") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">结束时间</td>
    <td bgcolor="#FFFFFF"><%= rs("终止时间") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">中断时间</td>
    <td bgcolor="#FFFFFF"><%= rs("中断时间") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">有效时间</td>
    <td bgcolor="#FFFFFF"><%= rs("合计时间")-rs("中断时间") %>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">不良现象</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF">不良总数</td>
    <td bgcolor="#FFFFFF"><%= rs("不良数") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">取数</td>
    <td bgcolor="#FFFFFF"><%= request.Cookies("qs") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">标准时间</td>
    <td bgcolor="#FFFFFF"><%= rs("标准时间") %>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">机台</td>
    <td bgcolor="#FFFFFF"><%= rs("机台") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">担当</td>
    <td bgcolor="#FFFFFF"><%= rs("管理担当") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">作业员</td>
    <td bgcolor="#FFFFFF"><%= rs("作业者") %>&nbsp;</td>
    <td bgcolor="#FFFFFF">生产效率</td>
    <td bgcolor="#FFFFFF"><%= rs("生产效率") %>&nbsp;</td>
  </tr>
  <tr>
    <td height="40" colspan="8" align="center" bgcolor="#FFFFFF" style="font-size:24px; font-weight:bold;"><a href="filshset.asp">确认生产完成</a></td>
  </tr>
</table>
</body>
<% 
call closeRs(rs)
call closeConn()
 %>
</html>
