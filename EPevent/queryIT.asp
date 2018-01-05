<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
id=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from EPevent where id='"& id &"'",1,3)
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
      <td align="center" bgcolor="#86C4C4"  style="color:#FFFFFF; font-weight:bold;">设备维护改善依赖申报</td>
    </tr>
  </table>
  <table width="95%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖人：</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("comuser") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">依赖部门</td>
      <td bgcolor="#FFFFFF"><%= rs("维护项目") %></td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">手机/短号：</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("phone") %><%= rs("tele") %></td>
      <td width="12%" align="right" bgcolor="#FFFFFF">期望纳期</td>
      <td bgcolor="#FFFFFF"><%= rs("nq")  %>&nbsp;</td>
    </tr>
	    <tr>
      <td width="15%" align="right" bgcolor="#FFFFFF">依赖时间：</td>
      <td width="20%" bgcolor="#FFFFFF"><%= rs("依赖时间") %></td>
      <td width="12%" rowspan="2" align="right" bgcolor="#FFFFFF">维护建议:</td>
      <td rowspan="2" bgcolor="#FFFFFF"><textarea name="jy" cols="40" rows="4" id="jy"></textarea>
        <input type="hidden" name="whid" id="whid" value="<%=id%>" />
        <input type="hidden" name="ipdress" id="ipdress" value="<%=ipdress%>"></td>
    </tr>

    <tr>
      <td align="right" bgcolor="#FFFFFF">紧急度</td>
      <td bgcolor="#FFFFFF"><%= rs("jj") %></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">依赖详情/现象描述：</td>
      <td bgcolor="#FFFFFF"><%= rs("现象描述") %>&nbsp;</td>
      <td align="right" bgcolor="#FFFFFF">维护人</td>
      <td bgcolor="#FFFFFF"><select name="whr" id="whr">
        <option value="0">请选择</option>
        <option value="2">周金山</option>
        <option value="3">钟祥林</option>
        <option value="4">李自伟</option>
	<option value="4">李德全</option>
	<option value="4">胡永平</option>
      </select>      </td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
      <td align="center"><label>
        <input name="Submit" type="submit" class="tableborder" onclick="return confirm('确认处理完成')" value="处 理 完 成" />
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






