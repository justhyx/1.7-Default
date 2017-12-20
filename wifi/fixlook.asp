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
<table width="80%" border="0" cellpadding="4" cellspacing="1" style="margin-left:30px" >
<tr><td colspan="5">&nbsp;</td></tr>
  <tr>
  <td width="13%"><a href="computer02.asp" target="frmright"><span style="font-weight:bold;">电脑购入记录</span></a></td>
  <td width="13%"><a href="computer03.asp" target="frmright"><span style="font-weight:bold;">电脑使用明细</span></a></td>
  <td width="13%"><a href="computer04.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">电脑维护记录</span></a></td>
  <td width="52%">&nbsp;</td>
<td width="9%" style="font-weight:bold;"algin="right"></td>
  </tr>
</table>
  <table width="95%" border="0" align="center" cellpadding="10" cellspacing="0" bordercolor="#000000">
    <tr>
      <td align="center"  style="font-weight:bold;">日常IT事务依赖申报</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="0"  bordercolor="#000000">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖人：</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("comuser") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖事项:</td>
      <td width="35%"  bgcolor="#FFFFFF"><%= rs("维护项目") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">手机/短号：</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("phone") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">内线电话:</td>
      <td width="35%"  bgcolor="#FFFFFF"><%= rs("tele") %></td>
    </tr>
	    <tr>
	 <td align="right" bgcolor="#FFFFFF">依赖原因：</td>
      <td  bgcolor="#FFFFFF"><%= rs("维护原因") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖时间：</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("依赖时间") %></td>

    </tr>
	    <tr>
		      <td width="15%" align="right" bgcolor="#FFFFFF">维护电脑:</td>
      <td  bgcolor="#FFFFFF"><%= rs("MAC") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">维护时间：</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("维护日期") %></td>

    </tr>
	    <tr>
			  <td width="15%" align="right" bgcolor="#FFFFFF">状态:</td>
      <td bgcolor="#FFFFFF"><%= rs("状态") %></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">维护人：</td>
      <td width="35%" bgcolor="#FFFFFF"><%= rs("维护人") %></td>

    </tr>

    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">IT意见:</td>
      <td bgcolor="#FFFFFF"><%= rs("IT意见") %></td>
	   <td align="right" bgcolor="#FFFFFF">备注：</td>
      <td bgcolor="#FFFFFF"><%= rs("备注") %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">依赖详情/现象描述：</td>
      <td colspan="3" bgcolor="#FFFFFF"><%= rs("现象描述") %>&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>	  <input type="button" name="Submit220" onclick="javascript:history.back();" value=" 返  回 " /></label>
	  <input type="button" name="Submit222" onclick="javascript:window.location.href='changetwo.asp?id=+ <%= rs("id") %> +';" value=" 修  改 " />
	    <input type="button" name="Submit222" onclick="javascript:window.location.href='comchk.asp?event=del03&id=+ <%= rs("id") %> +';" value=" 删  除 " />
	  </td>
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




