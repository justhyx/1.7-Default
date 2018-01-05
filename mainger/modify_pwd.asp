<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="checkLogin.asp?ee=modify">
  <table width="400" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
    <tr>
      <td colspan="2" align="center">密码修改</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">用户名：</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="uid" type="text" id="uid" />
      </label></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">原密码：</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="opwd" type="password" id="opwd" />
      </label></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">新密码：</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="npwd" type="password" id="npwd" />
      </label></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">确认密码</td>
      <td bgcolor="#FFFFFF"><label>
        <input name="kpwd" type="password" id="kpwd" />
      </label></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><label>
        <input type="submit" name="Submit" value="提 交" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
