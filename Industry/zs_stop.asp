<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ͳ��ǳϹ�ҵ����ϵͳ</title>
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
        <td><form id="form2" name="form2" method="post" action="chk.asp?xinxi=���&amp;event=ͣ��ԭ��">
            <label>
            <input type="submit" name="Submit2" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--���ͣ��--|��������ȷ��')" value="�� ��" />
            </label>
                    </form></td>
        <td><form id="form3" name="form3" method="post" action="chk.asp?xinxi=����&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit3" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--����ͣ��--|��������ȷ��')" value="�� ��" />
          </label>
                </form></td>
        <td><form id="form13" name="form13" method="post" action="chk.asp?xinxi=��е�ֵ���&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit13" style="width:210px; height:60px; font-size:30px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��е�ֵ���--|��������ȷ��')" value="��е�ֵ���" />
          </label>
                </form></td>
        <td><form id="form12" name="form12" method="post" action="chk.asp?xinxi=�����&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit12" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�����ͣ��--|��������ȷ��')" value="�����" />
          </label>
                </form></td>
		<td></td>
      </tr>
      <tr>
        <td><form id="form4" name="form4" method="post" action="chk.asp?xinxi=���ϸ���&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit4" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--���ϸ���--|��������ȷ��')" value="���ϸ���" />
          </label>
                </form></td>
        <td><form id="form5" name="form5" method="post" action="chk.asp?xinxi=��Ͳ����&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit5" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��Ͳ����--|��������ȷ��')" value="��Ͳ����" />
          </label>
                </form></td>
        <td><form id="form6" name="form6" method="post" action="chk.asp?xinxi=��Ͳ����&amp;event=ͣ��ԭ��">
          <label>
            <input type="submit" name="Submit6" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��Ͳ����--|��������ȷ��')" value="��Ͳ����" />
            </label>
        </form></td>
        <td><form id="form7" name="form7" method="post" action="chk.asp?xinxi=ģ�߲���&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit7" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--ģ�߲���--|��������ȷ��')" value="ģ�߲���" />
          </label>
                </form></td>
      </tr>
      <tr>
        <td><form id="form11" name="form11" method="post" action="chk.asp?xinxi=��ˮ��&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit11" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��ˮ��ͣ��--|��������ȷ��')" value="��ˮ��" />
          </label>
                </form></td>
        <td><form id="form8" name="form8" method="post" action="chk.asp?xinxi=��ģ&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit8" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��ģͣ��--|��������ȷ��')" value="�� ģ" />
          </label>
                </form></td>
        <td><form id="form9" name="form9" method="post" action="chk.asp?xinxi=�޻�&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit9" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�޻�ͣ��--|��������ȷ��')" value="�� ��" />
          </label>
                </form></td>
        <td><form id="form10" name="form10" method="post" action="chk.asp?xinxi=�����л�&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit10" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�����л�--|��������ȷ��')" value="�����л�" />
          </label>
                </form></td>
      </tr>
      <tr>
        <td><form id="form24" name="form24" method="post" action="chk.asp?xinxi=��ģ&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit24" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��ģͣ��--|��������ȷ��')" value="�� ģ" />
          </label>
                </form></td>
        <td><form id="form25" name="form25" method="post" action="chk.asp?xinxi=�ƻ�ͣ��&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit25" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�ƻ�ͣ��--|��������ȷ��')" value="�ƻ�ͣ��" />
          </label>
                </form></td>
        <td><form id="form26" name="form26" method="post" action="chk.asp?xinxi=�����&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit26" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�����ͣ��--|��������ȷ��')" value="�����" />
          </label>
                </form></td>
        <td>
        <form id="form1" name="form1" method="post" action="chk.asp?event=other_stop">
		  ������<input name="xinxi" type="text" style="width:150px; height:70px; font-size:24px;" id="xinxi" />
		     <label>
		     <input type="submit" name="Submit" value="ȷ��" />
		     </label>
        </form></td>
		        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
