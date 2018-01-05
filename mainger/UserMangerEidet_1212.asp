<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<% 
if session("Power")<>9998 then
response.end()
end if
dim action,UserID,UserPassword,UserName,UserSex,Department,Position,Phone,Tel,Power,addTime
action=Trim(Request.QueryString("action"))
if action="add" then
txttitle="新增用户"
UserID=""
UserPassword1=""
UserPassword=""
UserName=""
UserSex=""
Department=""
Position=""
Phone=""
Tel=""
Power=""
rsid=""
useraction="adduser"
end if
if action="modify" then
txttitle="修改用户管理"
call CreateRs(rs,"select * from HUDSON_User where id="&Trim(Request.QueryString("id")),1,1)
UserID=rs("UserID")
UserName=rs("UserName")
UserSex=rs("UserSex")
Department=rs("Department")
Position=rs("Position")
Phone=rs("Phone")
Tel=rs("Tel")
Power=rs("Power")
useraction="modify"
rsid=Trim(Request.QueryString("id"))
call closeRs(rs)
end if
 %>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="checkUser.asp?id=<%= rsid %>&action=<%= useraction %>">
  <p>&nbsp;</p>
  <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr>
      <td colspan="6" align="center" bgcolor="#86C4C4" style="font-weight:bold; color:#FFFFFF;"><%= txttitle %>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">厂牌号</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="UserID" type="text" id="UserID" value="<%= UserID %>" />
      </label></td>
      <td bgcolor="#FFFFFF">姓名</td>
      <td bgcolor="#FFFFFF"><input name="UserName" type="text" id="UserName" value="<%= UserName %>" /></td>
      <td bgcolor="#FFFFFF">职位</td>
      <td bgcolor="#FFFFFF"><input name="Position" type="text" id="Position" value="<%= Position %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">密　码</td>
      <td bgcolor="#FFFFFF"><input name="UserPassword" type="password" id="UserPassword" /></td>
      <td bgcolor="#FFFFFF">性别</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="UserSex" type="radio" value="男" checked="checked" />
      </label>
      男
      <input type="radio" name="UserSex" value="女" />
      女</td>
      <td bgcolor="#FFFFFF">手机号</td>
      <td bgcolor="#FFFFFF"><input name="Phone" type="text" id="Phone" value="<%= Phone %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">确认密码</td>
      <td bgcolor="#FFFFFF"><input name="UserPassword1" type="password" id="UserPassword1" /></td>
      <td bgcolor="#FFFFFF">部门</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="Department" id="Department">
          <option value="部品技术课">部品技术课</option>
          <option value="注塑检查课">注塑检查课</option>
		  <option value="生产管理课">生产管理课</option>
		  <option value="营业采购课">营业采购课</option>
		  <option value="人事总务课">人事总务课</option>
  		  <option value="仓库管理课">仓库管理课</option>
		  <option value="成形课">成形课</option>
		  <option value="技术课">技术课</option>
		  <option value="生产课">生产课</option>
		  <option value="设计组">设计组</option>
		  <option value="财务课">财务课</option>
        </select>
      </label></td>
      <td bgcolor="#FFFFFF">内线</td>
      <td bgcolor="#FFFFFF"><input name="Tel" type="text" id="Tel" value="<%= Tel %>" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">入厂时间</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="adddate" type="text" id="adddate" value="<%= date() %>" />
      </label></td>
      <td>组别</td>
      <td><select name="user_l" id="user_l">
        <option value="1">部品技术课</option>
        <option value="9">注塑检查课</option>
        <option value="6">生产管理课</option>
        <option value="8">营业采购课</option>
        <option value="4">人事总务课</option>
        <option value="3">技术课</option>
        <option value="7">生产课</option>
        <option value="5">设计组</option>
        <option value="2">财务课</option>
		<option value="20">无</option>
      </select></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">权 限</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="Power" id="Power">
          <option value="9999">普通用户</option>
          <option value="999">班长</option>
		  <option value="99">主管</option>
          <option value="9">经理</option>
          <option value="7">场长</option>
		  <option value="6">总经理</option>
          <option value="5">董事</option>
        </select>
      </label></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="确　定" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
<% call closeConn() %>