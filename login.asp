<%
	 dim rndnum,verifycode
	Randomize
	Do While Len(rndnum)<4
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	session("verifycode")=rndnum
%>
<html>
<head>

<title>和诚锴诚事务管理系统----登录</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css" type="text/css">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onmouseover="self.status='欢迎您的光临';return true" onselectstart="return false">
<form name="form1" method="post" action="mainger/checklogin.asp">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" background="img/leftxs.jpg">
    <tr>
      <td>&nbsp;</td>
      <td width="712" height="324" background="img/logon.jpg"><table width="100%" height="135" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="25%" height="18">&nbsp;</td>
          <td width="45%">&nbsp;</td>
          <td width="28%">&nbsp;</td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr>
          <td height="26">&nbsp;</td>
          <td>&nbsp;</td>
          <td><input name="UserID" type="text" id="UserID" size="18"></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="30">&nbsp;</td>
          <td align="right" style="color:#FFFFFF; font-weight:bold;">密　码&nbsp;&nbsp;&nbsp;</td>
          <td><label>
            <input name="pwd" type="password" id="pwd" size="18">
          </label></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="29">&nbsp;</td>
          <td>&nbsp;</td>
          <td align="center"><a href="mainger/modify_pwd.asp">修改密码</a></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td style="color:#FFFFFF;">默认密码是工号后6位请自行修改密码</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td><label>
            <input type="submit" name="Submit" value="登 录">
          </label></td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <div align="center"><br>
    <span class="STYLE1"><%= Trim(Request.QueryString("txtmage")) %></span>
    <INPUT type=hidden value=CheckLogin 
name=method>
  </div>
</FORM>
</body>
</html>