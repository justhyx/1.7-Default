<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>和诚锴诚工业管理系统</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<% call CreateRs(rs,"select * from 不良计算 where 部番='"&request.Cookies("bf")&"' and 时间='"&date()&"'",1,1)  %>
<body>
<table width="1024" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr>
    <td><!--#include file="head.asp" --></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="1024" border="0" cellspacing="1" cellpadding="4">
        <tr>
          <td><form id="form39" name="form39" method="post" action="chk.asp?event=zc_js">
            <label>
            <input type="submit" name="Submit39" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品做完一箱如果装箱请点击确定')" value="计 数" />
            </label>
                    </form></td>
          <td><table border="0" cellspacing="1" cellpadding="3">
            <tr>
              <td><%= request.Cookies("jhs") %></td>
            </tr>
            <tr>
              <td><%= left(cdbl(rs("计数"))/cdbl(request.cookies("zxs")),3) %></td>
            </tr>
            <tr>
              <td><%= rs("计数") %></td>
            </tr>
          </table></td>
          <td><form id="form3" name="form3" method="post" action="chk.asp?event=生产&amp;bottn=a">
              <label>
              <input type="submit" name="Submit2" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--黑点--|如果存在请点击确定')" value="黑 点" />
              </label>
              <%= rs("黑点") %>
            </form></td>
          <td><form id="form4" name="form4" method="post" action="chk.asp?event=生产&amp;bottn=c">
              <label>
              <input type="submit" name="Submit3" style="width:210px; height:60px; font-size:36px; font-weight:bold;" value="银 纹" onclick="return confirm('请确定部品存在|--银纹--|如果存在请点击确定')" />
              </label>
              <%= rs("银纹") %>
            </form></td>
          <td><form id="form5" name="form5" method="post" action="chk.asp?event=生产&amp;bottn=d">
              <label>
              <input type="submit" name="Submit4" style="width:210px; height:60px; font-size:36px; font-weight:bold;" value="黑 纹" onclick="return confirm('请确定部品存在|--黑纹--|如果存在请点击确定')" />
              </label>
              <%= rs("黑纹") %>
            </form></td>
        </tr>
        <tr>
          <td colspan="2"><form id="form1" name="form1" method="post" action="chk.asp?event=wxs">
              <label>
              <input name="wxs" type="text" id="wxs" />
              </label>
              <label>
              <input type="submit" name="Submit" value="尾箱数" />
              </label>
            </form></td>
          <td><form id="form6" name="form6" method="post" action="chk.asp?event=生产&amp;bottn=e">
              <label>
              <input type="submit" name="Submit5" style="width:210px; height:60px; font-size:36px; font-weight:bold;" value="缩 水" onclick="return confirm('请确定部品存在|--缩水--|如果存在请点击确定')" />
              </label>
              <%= rs("缩水") %>
            </form></td>
          <td><form id="form7" name="form7" method="post" action="chk.asp?event=生产&amp;bottn=f">
              <label>
              <input type="submit" name="Submit6" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--划伤--|如果存在请点击确定')" value="划 伤" />
              </label>
              <%= rs("划伤") %>
            </form></td>
          <td><form id="form8" name="form8" method="post" action="chk.asp?event=生产&amp;bottn=g">
              <label>
              <input type="submit" name="Submit7" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--毛刺--|如果存在请点击确定')" value="毛刺" />
              </label>
              <%= rs("毛刺") %>
            </form></td>
        </tr>
        <tr>
          <td colspan="2"><form id="form18" name="form18" method="post" action="chk.asp?event=生产统计">
            <label>
            <input type="submit" name="Submit17" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品生产完成？如果完成请点击确定')" value="完 成" />
            </label>
                              </form>            <a href="chk.asp?event=生产统计" onclick="return confirm('请确定部品生产完成？如果完成请点击确定')"></a></td>
          <td><form id="form9" name="form9" method="post" action="chk.asp?event=生产&amp;bottn=h">
              <label>
              <input type="submit" name="Submit8" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--气泡--|如果存在请点击确定')" value="气 泡" />
              </label>
              <%= rs("气泡") %>
            </form></td>
          <td><form id="form10" name="form10" method="post" action="chk.asp?event=生产&amp;bottn=i">
              <label>
              <input type="submit" name="Submit9" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--断裂--|如果存在请点击确定')" value="断 裂" />
              </label>
              <%= rs("断裂") %>
            </form></td>
          <td><form id="form11" name="form11" method="post" action="chk.asp?event=生产&amp;bottn=j">
              <label>
              <input type="submit" name="Submit10" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--顶白--|如果存在请点击确定')" value="顶 白" />
              </label>
              <%= rs("顶白") %>
            </form></td>
        </tr>
        <tr>
          <td height="30" colspan="2">&nbsp;</td>
          <td rowspan="2"><form id="form12" name="form12" method="post" action="chk.asp?event=生产&amp;bottn=k">
              <label>
              <input type="submit" name="Submit11" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--变形--|如果存在请点击确定')" value="变 形" />
              </label>
              <%= rs("变形") %>
            </form></td>
          <td rowspan="2"><form id="form13" name="form13" method="post" action="chk.asp?event=生产&amp;bottn=l">
              <label>
              <input type="submit" name="Submit12" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--碰伤--|如果存在请点击确定')" value="碰 伤" />
              </label>
              <%= rs("碰伤") %>
            </form></td>
          <td rowspan="2"><form id="form14" name="form14" method="post" action="chk.asp?event=生产&amp;bottn=b">
              <label>
              <input type="submit" name="Submit13" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--注塑不全--|如果存在请点击确定')" value="注塑不全" />
              </label>
              <%= rs("注塑不良") %>
            </form></td>
        </tr>
        <tr>
          <td colspan="2"><form id="form2" name="form2" method="post" action="chk.asp?event=cxqz_s">
              <label>
              <input name="cxzq_s" type="text" id="cxzq_s" />
              </label>
              <label>
              <input type="submit" name="Submit" value="实周期" />
              </label>
            </form>
            </td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
          <td><form id="form15" name="form15" method="post" action="chk.asp?event=生产&amp;bottn=m">
              <label>
              <input type="submit" name="Submit14" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--脱模拉伤--|如果存在请点击确定')" value="脱模拉伤" />
              </label>
              <%= rs("脱模拉伤") %>
            </form></td>
          <td><form id="form16" name="form16" method="post" action="chk.asp?event=生产&amp;bottn=n">
              <label>
              <input type="submit" name="Submit15" style="width:210px; height:60px; font-size:30px; font-weight:bold;" onclick="return confirm('请确定部品存在|--机械手跌落--|如果存在请点击确定')" value="机械手跌落" />
              </label>
              <%= rs("机械手跌落") %>
            </form></td>
          <td><form id="form17" name="form17" method="post" action="chk.asp?event=生产&amp;bottn=p">
              <label>
              <input type="submit" name="Submit16" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确定部品存在|--寸法不良--|如果存在请点击确定')" value="寸法不良" />
              </label>
              <%= rs("寸法不良") %>
            </form>
            <% call closeRs(rs) %></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<% call closeConn() %>
