<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
 <script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
<head>

<% 
id=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from 电脑维护记录 where id='"& id &"'",1,3)

 %>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form name="form01" id="form01" method="post" action="comchk.asp?event=add03">
  <table width="95%" border="0" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td align="center"  style="font-weight:bold;">维护记录添加</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖人：</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="fm" id="fm"></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖事项:</td>
      <td width="35%"  bgcolor="#FFFFFF"><input type="text" name="thing" id="thing"></td>
    </tr>
	    <tr>
	 <td align="right" bgcolor="#FFFFFF">依赖原因：</td>
      <td  bgcolor="#FFFFFF"><input type="text" name="reason" id="reason" size="50"></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖时间：</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="sd" id="sd" onClick="WdatePicker()" style="hand();"></td>

    </tr>
	    <tr>
		      <td width="15%" align="right" bgcolor="#FFFFFF">维护电脑:</td>
      <td  bgcolor="#FFFFFF"><input type="text" name="fc" id="fc"></td>
      <td width="15%" align="right" bgcolor="#FFFFFF">维护时间：</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="ft" id="ft" onClick="WdatePicker()" style="hand();"></td>

    </tr>
	    <tr>
			  <td width="15%" align="right" bgcolor="#FFFFFF">状态:</td>
              <td bgcolor="#FFFFFF">	  <select name="zt">
	  <option value="<%=rs("状态")%>" selected="selected"><%=rs("状态")%></option>
	  <option value="未处理">未处理</option>
	  <option value="已处理" selected="selected">已处理</option>
	  </select>
	  </td>
              <td width="15%" align="right" bgcolor="#FFFFFF">维护人：</td>
      <td width="35%" bgcolor="#FFFFFF"><input type="text" name="it" id="it"></td>

    </tr>

    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">IT意见:</td>
      <td bgcolor="#FFFFFF"><input type="text" name="im" id="im" size="50"></td>
	   <td align="right" bgcolor="#FFFFFF">备注：</td>
      <td bgcolor="#FFFFFF"><input type="text" name="bz" id="bz"></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">依赖详情/现象描述：</td>
      <td colspan="3" bgcolor="#FFFFFF"><input type="text" name="xx" id="xx" size="100">&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>	  <input type="submit" name="Submit220"  value=" 添  加 " /></label></td>
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




