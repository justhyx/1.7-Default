<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #00CCCC}
-->
</style>
</head>
<%  call CreateRs(rs,"select * from meeting where id="&Trim(Request.QueryString("id")),1,1) %>
<body>
<p>&nbsp;</p>
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr>
    <td colspan="4" align="center" style="font-weight:bold;">会议室使用申请表</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>申请部门</td>
    <td><span class="STYLE1"><%= rs("M_Department") %>&nbsp;</span></td>
    <td>申请人</td>
    <td><span class="STYLE1"><%= rs("M_usename") %></span></td>
  </tr>
  <tr>
    <td>使用时间</td>
    <td><span class="STYLE1"><%= rs("M_usetime") %>&nbsp;&nbsp;<%= rs("M_usetime_h") %>&nbsp;&nbsp;时</span></td>
	<td>结束时间</td>
    <td><span class="STYLE1"><%= rs("M_overtime") %>&nbsp;&nbsp;<%= rs("M_overtime_h") %>&nbsp;&nbsp;时</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>客户名称</td>
    <td><span class="STYLE1"><%= rs("M_guest") %></span></td>
    <td>预计人数</td>
    <td><span class="STYLE1"><%= rs("M_number") %></span></td>
  </tr>
  <tr>
    <td>设备需求</td>
    <td><span class="STYLE1"><%= rs("M_device") %>&nbsp;&nbsp;<%= rs("M_device1") %></span></td>
    <td>使用会议室</td>
    <td><span class="STYLE1"><%= rs("M_meetingroom") %>号会议室</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>会议室状态</td>
    <td><span class="STYLE1"><%= rs("M_flog") %>&nbsp;</span></td>
    <td>管理者</td>
    <td><span class="STYLE1"><%= rs("M_manger") %></span></td>
  </tr>
  <tr>
    <td >提交时间</td>
    <td ><span class="STYLE1"><%= rs("addtime") %> </span></td>
    <td >完成时间</td>
    <td ><span class="STYLE1"><%= rs("uptime") %></span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td >意向会议室</td>
    <td ><%=  rs("intention")%>&nbsp;</td>
    <td >&nbsp;</td>
    <td >&nbsp;</td>
  </tr>
</table>
<%
if session("Department")="行政部" then
%>
<form id="form1" name="form1" method="post" action="chkmeeting.asp">
  <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="11%">会议室安排状态</td>
      <td width="89%"><label>
        <select name="M_flog" id="M_flog">
          <option value="预订完成">预订完成</option>
		  <option value="使用中">使用中</option>          
		  <option value="使用结束">使用结束</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td>会议室号</td>
      <td><label>
        <select name="M_meetingroom" id="M_meetingroom">
          <option value="1">1号会议室</option>
          <option value="2">2号会议室</option>
          <option value="3">3号会议室</option>
          <option value="4">4号会议室</option>
          <option value="5">5号会议室</option>
          <option value="6">6号会议室</option>
          <option value="7">7号会议室</option>
		  <option value="8">8号会议室</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td><label>
        <input type="submit" name="Submit" value="提　交" />
      </label></td>
      <td><input name="action" type="hidden" id="action" value="update" />
      <input name="id" type="hidden" id="id" value="<%= rs("id") %>" /></td>
    </tr>
  </table>
  <!--#include file="mettuse.asp" -->
</form>
<% 
end if
call closeRs(rs)
call closeConn()
 %>
</body>
</html>







