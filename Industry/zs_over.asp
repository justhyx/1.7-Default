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
    <td><table width="1024" border="0" cellspacing="1" cellpadding="4">
      <tr>
        <td><form id="form1" name="form1" method="post" action="chk.asp?xinxi=��ģ&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��ģͣ��--|��������ȷ��')" value="�� ģ" />
          </label>
                </form></td>
        <td><form id="form2" name="form2" method="post" action="chk.asp?xinxi=��ɫ&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit2" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��ɫͣ��--|��������ȷ��')" value="�� ɫ" />
          </label>
                </form></td>
        <td><form id="form6" name="form6" method="post" action="chk.asp?xinxi=�����&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit6" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�����ͣ��--|��������ȷ��')" value="��ģ��ɫ" />
          </label>
                        </form>
        </td>
        </tr>
      <tr>
        <td><form id="form4" name="form4" method="post" action="chk.asp?xinxi=��ģ&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit4" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--��ģͣ��--|��������ȷ��')" value="�� ģ" />
          </label>
                </form></td>
        <td><form id="form5" name="form5" method="post" action="chk.asp?xinxi=�ƻ�ͣ��&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit5" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�ƻ�ͣ��--|��������ȷ��')" value="�ƻ�ͣ��" />
          </label>
                </form></td>
        <td><form id="form3" name="form3" method="post" action="chk.asp?xinxi=�����&amp;event=ͣ��ԭ��">
          <label>
          <input type="submit" name="Submit3" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ĳ�����|--�����ͣ��--|��������ȷ��')" value="�����" />
          </label>
                </form></td>
        </tr>
      
    </table></td>
  </tr>
</table>
</body>
</html>
