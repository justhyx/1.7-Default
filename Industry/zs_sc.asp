<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ͳ��ǳϹ�ҵ����ϵͳ</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<% call CreateRs(rs,"select * from �������� where ����='"&request.Cookies("bf")&"' and ʱ��='"&date()&"'",1,1)  %>
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
            <input type="submit" name="Submit39" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����һ�����װ������ȷ��')" value="�� ��" />
            </label>
                    </form></td>
          <td><table border="0" cellspacing="1" cellpadding="3">
            <tr>
              <td><%= request.Cookies("jhs") %></td>
            </tr>
            <tr>
              <td><%= left(cdbl(rs("����"))/cdbl(request.cookies("zxs")),3) %></td>
            </tr>
            <tr>
              <td><%= rs("����") %></td>
            </tr>
          </table></td>
          <td><form id="form3" name="form3" method="post" action="chk.asp?event=����&amp;bottn=a">
              <label>
              <input type="submit" name="Submit2" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--�ڵ�--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("�ڵ�") %>
            </form></td>
          <td><form id="form4" name="form4" method="post" action="chk.asp?event=����&amp;bottn=c">
              <label>
              <input type="submit" name="Submit3" style="width:210px; height:60px; font-size:36px; font-weight:bold;" value="�� ��" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" />
              </label>
              <%= rs("����") %>
            </form></td>
          <td><form id="form5" name="form5" method="post" action="chk.asp?event=����&amp;bottn=d">
              <label>
              <input type="submit" name="Submit4" style="width:210px; height:60px; font-size:36px; font-weight:bold;" value="�� ��" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" />
              </label>
              <%= rs("����") %>
            </form></td>
        </tr>
        <tr>
          <td colspan="2"><form id="form1" name="form1" method="post" action="chk.asp?event=wxs">
              <label>
              <input name="wxs" type="text" id="wxs" />
              </label>
              <label>
              <input type="submit" name="Submit" value="β����" />
              </label>
            </form></td>
          <td><form id="form6" name="form6" method="post" action="chk.asp?event=����&amp;bottn=e">
              <label>
              <input type="submit" name="Submit5" style="width:210px; height:60px; font-size:36px; font-weight:bold;" value="�� ˮ" onclick="return confirm('��ȷ����Ʒ����|--��ˮ--|�����������ȷ��')" />
              </label>
              <%= rs("��ˮ") %>
            </form></td>
          <td><form id="form7" name="form7" method="post" action="chk.asp?event=����&amp;bottn=f">
              <label>
              <input type="submit" name="Submit6" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("����") %>
            </form></td>
          <td><form id="form8" name="form8" method="post" action="chk.asp?event=����&amp;bottn=g">
              <label>
              <input type="submit" name="Submit7" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--ë��--|�����������ȷ��')" value="ë��" />
              </label>
              <%= rs("ë��") %>
            </form></td>
        </tr>
        <tr>
          <td colspan="2"><form id="form18" name="form18" method="post" action="chk.asp?event=����ͳ��">
            <label>
            <input type="submit" name="Submit17" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ������ɣ�����������ȷ��')" value="�� ��" />
            </label>
                              </form>            <a href="chk.asp?event=����ͳ��" onclick="return confirm('��ȷ����Ʒ������ɣ�����������ȷ��')"></a></td>
          <td><form id="form9" name="form9" method="post" action="chk.asp?event=����&amp;bottn=h">
              <label>
              <input type="submit" name="Submit8" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("����") %>
            </form></td>
          <td><form id="form10" name="form10" method="post" action="chk.asp?event=����&amp;bottn=i">
              <label>
              <input type="submit" name="Submit9" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("����") %>
            </form></td>
          <td><form id="form11" name="form11" method="post" action="chk.asp?event=����&amp;bottn=j">
              <label>
              <input type="submit" name="Submit10" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("����") %>
            </form></td>
        </tr>
        <tr>
          <td height="30" colspan="2">&nbsp;</td>
          <td rowspan="2"><form id="form12" name="form12" method="post" action="chk.asp?event=����&amp;bottn=k">
              <label>
              <input type="submit" name="Submit11" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("����") %>
            </form></td>
          <td rowspan="2"><form id="form13" name="form13" method="post" action="chk.asp?event=����&amp;bottn=l">
              <label>
              <input type="submit" name="Submit12" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')" value="�� ��" />
              </label>
              <%= rs("����") %>
            </form></td>
          <td rowspan="2"><form id="form14" name="form14" method="post" action="chk.asp?event=����&amp;bottn=b">
              <label>
              <input type="submit" name="Submit13" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--ע�ܲ�ȫ--|�����������ȷ��')" value="ע�ܲ�ȫ" />
              </label>
              <%= rs("ע�ܲ���") %>
            </form></td>
        </tr>
        <tr>
          <td colspan="2"><form id="form2" name="form2" method="post" action="chk.asp?event=cxqz_s">
              <label>
              <input name="cxzq_s" type="text" id="cxzq_s" />
              </label>
              <label>
              <input type="submit" name="Submit" value="ʵ����" />
              </label>
            </form>
            </td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
          <td><form id="form15" name="form15" method="post" action="chk.asp?event=����&amp;bottn=m">
              <label>
              <input type="submit" name="Submit14" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--��ģ����--|�����������ȷ��')" value="��ģ����" />
              </label>
              <%= rs("��ģ����") %>
            </form></td>
          <td><form id="form16" name="form16" method="post" action="chk.asp?event=����&amp;bottn=n">
              <label>
              <input type="submit" name="Submit15" style="width:210px; height:60px; font-size:30px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--��е�ֵ���--|�����������ȷ��')" value="��е�ֵ���" />
              </label>
              <%= rs("��е�ֵ���") %>
            </form></td>
          <td><form id="form17" name="form17" method="post" action="chk.asp?event=����&amp;bottn=p">
              <label>
              <input type="submit" name="Submit16" style="width:210px; height:60px; font-size:36px; font-weight:bold;" onclick="return confirm('��ȷ����Ʒ����|--�編����--|�����������ȷ��')" value="�編����" />
              </label>
              <%= rs("�編����") %>
            </form>
            <% call closeRs(rs) %></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<% call closeConn() %>
