<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% call CreateRs(rs,"select * from �������� where ����='"&request.Cookies("bf_one")&"' and ʱ��='"&date()&"' and ��̨='"&request.cookies("scjt")&"'",1,1)  %>
</head>
<body>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="500" style="font-weight:bold;">��Ʒ���ţ�<%= request.cookies("bf_one") %></td>
  </tr>
</table>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="125" height="70"><a href="chk_two.asp?event=zc_js"  onclick="return confirm('��ȷ�����һ����������������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td width="125" height="70"><a href="chk_two.asp?event=����&amp;bottn=a" onclick="return confirm('��ȷ����Ʒ����|--�ڵ�--|�����������ȷ��')"><img src="../img/dblshow_rd/�ڵ�.gif" width="123" height="60" border="none" /></a></td>
    <td width="125" height="70"><a href="chk_two.asp?event=����&amp;bottn=c" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td width="125" height="70"><a href="chk_two.asp?event=����&amp;bottn=d" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125" height="70"><form id="form1" name="form1" method="post" action="chk_two.asp?event=wxs">����������<%= rs("����")%>�䣺<%= left(cdbl(rs("����"))/cdbl(request.cookies("zxs_one")),3)  %><br/>     
      <input name="wxs" type="text" id="wxs" size="5" />    
        <input type="submit" name="Submit" value="β����" />
    </form>    </td>
    <td><a href="chk_two.asp?event=����&amp;bottn=e" onclick="return confirm('��ȷ����Ʒ����|--��ˮ--|�����������ȷ��')"><img src="../img/dblshow_rd/��ˮ.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=f" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=g" onclick="return confirm('��ȷ����Ʒ����|--ë��--|�����������ȷ��')"><img src="../img/dblshow_rd/ë��.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125" height="70"><a href="chk_two.asp?event=����ͳ��" onclick="return confirm('��ȷ����Ʒ������ɣ�����������ȷ��')"><img src="../img/dblshow_rd/���.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=h" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=i" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=j" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125"><form id="form2" name="form2" method="post" action="chk_two.asp?event=cxzq_s">
	 <input name="cxzq_s" type="text" id="cxzq_s" size="5" />    
        <input type="submit" name="Submit" value="ʵ����" />
    </form></td>
    <td height="70"><a href="chk_two.asp?event=����&amp;bottn=k" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=l" onclick="return confirm('��ȷ����Ʒ����|--����--|�����������ȷ��')"><img src="../img/dblshow_rd/����.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=b" onclick="return confirm('��ȷ����Ʒ����|--ע�ܲ�ȫ--|�����������ȷ��')"><img src="../img/dblshow_rd/ע�ܲ�ȫ.gif" width="123" height="60" border="none" /></a></td>
  </tr>
  <tr>
    <td width="125"><table width="125" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000">
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("�ڵ�") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("��ˮ") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("ë��") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("ע�ܲ���") %>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><%= rs("��ģ����") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("��е�ֵ���") %>&nbsp;</td>
        <td bgcolor="#FFFFFF"><%= rs("�編����") %>&nbsp;</td>
      </tr>
    </table></td>
    <td height="70"><a href="chk_two.asp?event=����&amp;bottn=m" onclick="return confirm('��ȷ����Ʒ����|--��ģ����--|�����������ȷ��')"><img src="../img/dblshow_rd/��ģ����.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=n" onclick="return confirm('��ȷ����Ʒ����|--��е�ֵ���--|�����������ȷ��')"><img src="../img/dblshow_rd/��е������.gif" width="123" height="60" border="none" /></a></td>
    <td><a href="chk_two.asp?event=����&amp;bottn=p" onclick="return confirm('��ȷ����Ʒ����|--�編����--|�����������ȷ��')"><img src="../img/dblshow_rd/�編����.gif" width="123" height="60" border="none" /></a></td>
  </tr>
</table>
</body>
<% 
call closeRs(rs)
call closeConn()
 %>
</html>
