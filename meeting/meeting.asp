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
<form id="form1" name="form1" method="post" action="chkmeeting.asp">  

<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr>
      <td colspan="4" align="center" style="font-weight:bold;">会议室使用申请表</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>申请部门</td>
      <td><label>
        <input name="M_Department" type="text" id="M_Department" value="<%= session("Department")%>" />
      </label></td>
      <td>申请人</td>
      <td><input name="M_usename" type="text" id="M_usename" value="<%= session("UserName")%>" /></td>
    </tr>
    <tr>
      <td>使用时间</td>
      <td><input name="M_usetime" type="text" id="M_usetime" onClick="WdatePicker()" style="hand();" />
        <img onclick="WdatePicker({el:$dp.$('StartTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">&nbsp;
        <label>
        <input name="M_usetime_h" type="text" id="M_usetime_h" size="4" />
        时        </label></td>
      <td>结束时间</td>
      <td><input name="M_overtime" type="text" id="M_overtime" onClick="WdatePicker()" />
      <img onclick="WdatePicker({el:$dp.$('OverTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">&nbsp;
      <label>
<input name="M_overtime_h" type="text" id="M_overtime_h" size="4" /> 
时      </label></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>客户名称</td>
      <td><input name="M_guest" type="text" id="M_guest" /></td>
      <td>预计人数</td>
      <td><input name="M_number" type="text" id="M_number" /></td>
    </tr>
    <tr>
      <td>设备需求</td>
      <td colspan="3"><label>
        <input name="M_device" type="checkbox" id="M_device" value="投影仪" />
      投影仪 
      <input name="M_device1" type="checkbox" id="M_device1" value="电子白板" />
      电子白板</label></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>意向会议室</td>
      <td><label>
        <input name="intention" type="text" id="intention" size="8" />
      号会议室</label></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center"><label>
        <input type="submit" name="Submit" value="提　交" />
        <input name="action" type="hidden" id="action" value="add" />
      </label></td>
    </tr>
  </table>
</form>
  <!--#include file="mettuse.asp" -->
</body>
</html>
