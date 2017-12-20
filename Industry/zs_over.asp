<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>和诚锴诚工业管理系统</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
dim stop_ip
stop_ip=Request.ServerVariables("REMOTE_ADDR")
stop_time=now()
 %>
</head>

<body>
<table width="1024" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr>
    <td><table width="1024" border="0" cellspacing="1" cellpadding="4">
      <tr>
        <td><form id="form1" name="form1" method="post" action="chk.asp?xinxi=换模&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换模停机--|如果是请点确定')" value="换 模" />
          </label>
                </form></td>
        <td><form id="form2" name="form2" method="post" action="chk.asp?xinxi=换色&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit2" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换色停机--|如果是请点确定')" value="换 色" />
          </label>
                </form></td>
        <td><form id="form6" name="form6" method="post" action="chk.asp?xinxi=换镶件&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit6" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换镶件停机--|如果是请点确定')" value="换模换色" />
          </label>
                        </form>
        </td>
        </tr>
      <tr>
        <td><form id="form4" name="form4" method="post" action="chk.asp?xinxi=试模&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit4" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--试模停机--|如果是请点确定')" value="试 模" />
          </label>
                </form></td>
        <td><form id="form5" name="form5" method="post" action="chk.asp?xinxi=计划停机&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit5" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--计划停机--|如果是请点确定')" value="计划停机" />
          </label>
                </form></td>
        <td><form id="form3" name="form3" method="post" action="chk.asp?xinxi=换镶件&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit3" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换镶件停机--|如果是请点确定')" value="换镶件" />
          </label>
                </form></td>
        </tr>
      
    </table></td>
  </tr>
</table>
</body>
</html>
