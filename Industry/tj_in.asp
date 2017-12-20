<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="chk.asp?event=tj_in">
  <table width="101%" height="300" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="300" align="center" valign="middle"><label>
        <input type="submit" name="Submit" style="width:250px; height:140px; font-size:36px; font-weight:bold;" value="开始作业" />
      &nbsp;&nbsp;
      </label></td>
	  <td>
	  <form id="form2" name="form2" method="post" action="chk.asp?xinxi=装模异常&amp;event=停机原因">
		 <label>
		 <input type="button" name="Submit2" style="width:250px; height:140px; background-color:#FF0000; color:#FFFFFF; font-size:36px; font-weight:bold;" onclick="location.href= 'Industry/index.asp '" value="调机异常" /></label></td>
    </tr>
  </table>
</form>
</body>
</html>
