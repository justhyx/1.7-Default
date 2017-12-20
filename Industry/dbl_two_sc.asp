<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% call CreateRs(rs,"select * from 不良计算 where 部番='"&request.Cookies("bf_one")&"' and 时间='"&date()&"' and 机台='"&request.cookies("scjt")&"'",1,1)  %>
</head>
<body>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="500" style="font-weight:bold;">部品番号：<%= request.cookies("bf_one") %></td>
  </tr>
</table>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="125" height="70"><a href="chk_two.asp?event=zc_js"  onclick="return confirm('请确定完成一箱数量如果完成请点击确定')"><img src="../img/dblshow_rd/计数.gif" width="123" height="60" border="none" /></a></td>
    <td width="125" height="70"><a href="chk_two.asp?event=生产&amp;bottn=a" onclick="return confirm('请确定部品存在|--黑点--|如果存在请点击确定')"><img src="../img/dblshow_rd/黑点.gif" width="123" height="60" border="none" /></a></td>
    <td width="125" height="70"><a href="chk_two.asp?event=生产&amp;bottn=c" onclick="return confirm('请确定部品存在|--银纹--|如果存在请点击确定')"><img src="../img/dblshow_rd/银纹.gif" width="123" height="60" border="none" /></a></td>
    <td width="125" height="70"><a href="chk_two.asp?event=生产&amp;bottn=d" onclick="return confirm('请确定部品存在|--黑纹--|如果存在请点击确定')"><img src="../img/dblshow_rd/黑纹.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125" height="70"><form id="form1" name="form1" method="post" action="chk_two.asp?event=wxs">生产数量：<%= rs("计数")%>箱：<%= left(cdbl(rs("计数"))/cdbl(request.cookies("zxs_one")),3)  %><br/>     
      <input name="wxs" type="text" id="wxs" size="5" />    
        <input type="submit" name="Submit" value="尾箱数" />
    </form>    </td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=e" onclick="return confirm('请确定部品存在|--缩水--|如果存在请点击确定')"><img src="../img/dblshow_rd/缩水.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=f" onclick="return confirm('请确定部品存在|--划伤--|如果存在请点击确定')"><img src="../img/dblshow_rd/划伤.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=g" onclick="return confirm('请确定部品存在|--毛刺--|如果存在请点击确定')"><img src="../img/dblshow_rd/毛刺.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125" height="70"><a href="chk_two.asp?event=生产统计" onclick="return confirm('请确定部品生产完成？如果完成请点击确定')"><img src="../img/dblshow_rd/完成.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=h" onclick="return confirm('请确定部品存在|--气泡--|如果存在请点击确定')"><img src="../img/dblshow_rd/气泡.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=i" onclick="return confirm('请确定部品存在|--断裂--|如果存在请点击确定')"><img src="../img/dblshow_rd/断裂.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=j" onclick="return confirm('请确定部品存在|--顶白--|如果存在请点击确定')"><img src="../img/dblshow_rd/顶白.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125"><form id="form2" name="form2" method="post" action="chk_two.asp?event=cxzq_s">
	 <input name="cxzq_s" type="text" id="cxzq_s" size="5" />    
        <input type="submit" name="Submit" value="实周期" />
    </form></td>
    <td height="70"><a href="chk_two.asp?event=生产&amp;bottn=k" onclick="return confirm('请确定部品存在|--变形--|如果存在请点击确定')"><img src="../img/dblshow_rd/变形.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=l" onclick="return confirm('请确定部品存在|--碰伤--|如果存在请点击确定')"><img src="../img/dblshow_rd/碰伤.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=b" onclick="return confirm('请确定部品存在|--注塑不全--|如果存在请点击确定')"><img src="../img/dblshow_rd/注塑不全.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125"><table width="125" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000">
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("黑点") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("银纹") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("黑纹") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("缩水") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("划伤") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("毛刺") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("气泡") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("断裂") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("顶白") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("变形") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("碰伤") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("注塑不良") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("脱模拉伤") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("机械手跌落") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("寸法不良") %>&nbsp;</td>
      </tr>
    </table></td>
    <td height="70"><a href="chk_two.asp?event=生产&amp;bottn=m" onclick="return confirm('请确定部品存在|--脱模拉伤--|如果存在请点击确定')"><img src="../img/dblshow_rd/脱模拉伤.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=n" onclick="return confirm('请确定部品存在|--机械手跌落--|如果存在请点击确定')"><img src="../img/dblshow_rd/机械手脱落.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=生产&amp;bottn=p" onclick="return confirm('请确定部品存在|--寸法不良--|如果存在请点击确定')"><img src="../img/dblshow_rd/寸法不良.gif" width="123" height="60" border="none" /></a></td>
  </tr>
</table>
</body>
<% 
call closeRs(rs)
call closeConn()
 %>
</html>
