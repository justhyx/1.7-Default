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
    <td><!--#include file="head.asp" --></td>
  </tr>
  <tr>
    <td><hr align="left" width="650" size="1" noshade="noshade" /></td>
  </tr>
  <tr>
    <td><table width="1024" border="0" cellspacing="1" cellpadding="3">
      <tr>
        <td><form id="form2" name="form2" method="post" action="chk.asp?xinxi=点检&amp;event=停机原因">
            <label>
            <input type="submit" name="Submit2" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--点检停机--|如果是请点确定')" value="点 检" />
            </label>
                    </form></td>
        <td><form id="form3" name="form3" method="post" action="chk.asp?xinxi=调机&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit3" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--调机停机--|如果是请点确定')" value="调 机" />
          </label>
                </form></td>
        <td><form id="form13" name="form13" method="post" action="chk.asp?xinxi=机械手调整&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit13" style="width:210px; height:60px; font-size:30px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--机械手调整--|如果是请点确定')" value="机械手调整" />
          </label>
                </form></td>
        <td><form id="form12" name="form12" method="post" action="chk.asp?xinxi=换镶件&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit12" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换镶件停机--|如果是请点确定')" value="换镶件" />
          </label>
                </form></td>
		<td></td>
      </tr>
      <tr>
        <td><form id="form4" name="form4" method="post" action="chk.asp?xinxi=材料干燥&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit4" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--材料干燥--|如果是请点确定')" value="材料干燥" />
          </label>
                </form></td>
        <td><form id="form5" name="form5" method="post" action="chk.asp?xinxi=料筒降温&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit5" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--料筒降温--|如果是请点确定')" value="料筒降温" />
          </label>
                </form></td>
        <td><form id="form6" name="form6" method="post" action="chk.asp?xinxi=料筒升温&amp;event=停机原因">
          <label>
            <input type="submit" name="Submit6" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--料筒升温--|如果是请点确定')" value="料筒升温" />
            </label>
        </form></td>
        <td><form id="form7" name="form7" method="post" action="chk.asp?xinxi=模具擦油&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit7" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--模具擦油--|如果是请点确定')" value="模具擦油" />
          </label>
                </form></td>
      </tr>
      <tr>
        <td><form id="form11" name="form11" method="post" action="chk.asp?xinxi=换水咀&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit11" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换水咀停机--|如果是请点确定')" value="换水咀" />
          </label>
                </form></td>
        <td><form id="form8" name="form8" method="post" action="chk.asp?xinxi=修模&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit8" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--修模停机--|如果是请点确定')" value="修 模" />
          </label>
                </form></td>
        <td><form id="form9" name="form9" method="post" action="chk.asp?xinxi=修机&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit9" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--修机停机--|如果是请点确定')" value="修 机" />
          </label>
                </form></td>
        <td><form id="form10" name="form10" method="post" action="chk.asp?xinxi=电力切换&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit10" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--电力切换--|如果是请点确定')" value="电力切换" />
          </label>
                </form></td>
      </tr>
      <tr>
        <td><form id="form24" name="form24" method="post" action="chk.asp?xinxi=试模&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit24" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--试模停机--|如果是请点确定')" value="试 模" />
          </label>
                </form></td>
        <td><form id="form25" name="form25" method="post" action="chk.asp?xinxi=计划停机&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit25" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--计划停机--|如果是请点确定')" value="计划停机" />
          </label>
                </form></td>
        <td><form id="form26" name="form26" method="post" action="chk.asp?xinxi=换镶件&amp;event=停机原因">
          <label>
          <input type="submit" name="Submit26" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('请确认你的操作是|--换镶件停机--|如果是请点确定')" value="换镶件" />
          </label>
                </form></td>
        <td>
        <form id="form1" name="form1" method="post" action="chk.asp?event=other_stop">
		  其他：<input name="xinxi" type="text" style="width:150px; height:70px; font-size:24px;" id="xinxi" />
		     <label>
		     <input type="submit" name="Submit" value="确认" />
		     </label>
        </form></td>
		        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
