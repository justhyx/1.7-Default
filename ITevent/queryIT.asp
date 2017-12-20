<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
id=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from 电脑维护记录 where id='"& id &"'",1,3)
ipdress=request.ServerVariables("REMOTE_ADDR")

 %>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<p>&nbsp;</p>
<form id="form1" name="form1" method="post" action="checkIT.asp?action=IT_query">  
  <table width="95%" border="1" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td align="center" bgcolor="#86C4C4"  style="color:#FFFFFF; font-weight:bold;">日常IT事务依赖申报</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖人：</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("comuser") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">依赖事项:</td>
      <td colspan="2" bgcolor="#FFFFFF"><%= rs("维护项目") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">手机/短号：</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("phone") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">内线电话:</td>
      <td colspan="2" bgcolor="#FFFFFF"><%= rs("tele") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖时间：</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("依赖时间") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">IT意见:</td>
      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="itmind" id="itmind" size="50">
	  <input type="hidden" name="whid" id="whid" value="<%=id%>">
	  <input type="hidden" name="ipdress" id="ipdress" value="<%=ipdress%>"></td>
    </tr>

    <tr>
      <td align="right" bgcolor="#FFFFFF">依赖原因：</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= rs("维护原因") %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">依赖详情/现象描述：</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= rs("现象描述") %>&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>
        <input name="Submit" type="submit" class="tableborder" value="确 认 依 赖" />
      </label></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
<% 
call closeRs(rs)
call closeConn()
 %>






