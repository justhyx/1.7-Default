<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ģ���絥����ϵͳ---��ӡ</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT>
</OBJECT>
</head>
 <style media="print">

    .noprint { display: none }

  </style>
<body>
<table width="1050" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000" id="mytable">
  <tr style="font-weight:bold;">    
    <td height="60" colspan="13"><table width="1050" height="60" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2" align="center" style="font-size:24px; font-weight:bold; font-family:'����'">��ģ���絥</td>
      </tr>
      <tr>
        <td align="left"><%= date() %>&nbsp;
		<label class="noprint">
		          <form id="form1" name="form1" method="post" action="">
            <label>
            <input name="TxtKey" type="text" id="TxtKey" class="noprint" />
            </label>
            <label>
            <input type="submit" name="Submit" value="�� ѯ" class="noprint" />
            </label>
            <label class="noprint">�������ڲ�ѯ ����2012-6-7 ע�����������</label>
          </form>
		</label>		</td>
        <td width="200" align="left">���ϣ�</td>
      </tr>
    </table></td>
  </tr>
  <tr style="font-weight:bold;">    
    <td width="90">��/Ʒ��</td>
    <td width="40">��̨</td>
    <td width="60">OK��̨</td>
    <td width="40">Ѩ��</td>
    <td width="70">���ϱ��</td>
    <td width="40">����</td>
    <td width="40">����</td>
    <td width="40">PO</td>
    <td width="40">����</td>
    <td width="60"><a href="print.asp">����</a></td>
    <td width="60">��ע</td>
    <td width="60">ʱ�䰲��</td>
    <td width="60">����˵��</td>
  </tr>
   <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=999
'�����ܼ�¼������ҳ������ǰҳ��������������
Dim ThisRsCount,ThisPageCount,ThisCurrentPage,i
'��URL�л�ȡҳ��Page�����ж���Ч��
ThisCurrentPage=Request("Page")
If Request("Page")="" Then
	ThisCurrentPage=1
ElseIf IsNumeric(ThisCurrentPage) Then
	ThisCurrentPage=int(ThisCurrentPage)
Else
	ThisCurrentPage=1
End If
If ThisCurrentPage<1 Then ThisCurrentPage=1
%>
  <%
'�����ѯ�ؼ��ֱ�������URL����л�ȡ��ѯ�Ĺؼ���TxtKey������ϳ�ģ����ѯSQL���
Dim TxtKey
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from trymolde where addtime='"&date()&"' and okfolg='"&int(0)&"'Order By ���ϱ�� DESC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where addtime like '%"&TxtKey&"%' Order By ���ϱ�� DESC",1,1)
		if rs.eof or rs.bof then
			response.Write("û�д˼�¼")
			response.End()
		end if
End If	
%>
  <%
'��ȡ�ܼ�¼��
ThisRsCount=Rs.RecordCount
'����ÿҳ��ʾ��������¼
Rs.PageSize=ThisPageSize
'��ȡ��ҳ��
ThisPageCount=Rs.PageCount
'�жϵ�ǰҳ�����Ч��
If ThisCurrentPage>ThisPageCount Then ThisCurrentPage=ThisPageCount
'���õ�ǰҳ��
Rs.AbsolutePage=ThisCurrentPage
%>
  <%
'ѭ����ȡ����
i=0
Do While Not Rs.EOF and i<ThisPageSize
%>
  <tr>
    <td><%= rs("Ʒ��") %> | <%= rs("����") %></td>
    <td><%= rs("��׼Ҫ����̨") %></td>
    <td><%= rs("����ȷ�ϻ�̨") %></td>
    <td><%= rs("Ѩ��") %></td>
    <td><%= rs("���ϱ��") %></td>
    <td><%= rs("����") %></td>
    <td><%= rs("��������") %></td>
    <td><%= rs("PO") %></td>
    <td><%= rs("����") %></td>
    <td><%= rs("ģ�ߵ���") %></td>
    <td><%= rs("��ע") %></td>
    <td><%= rs("����ʱ��") %>&nbsp;</td>
    <td><%= rs("���ܱ�ע") %>&nbsp;</td>
  </tr>
  <%
	i=i+1
	Rs.MoveNext
Loop	
%>
</table>
   <center class=noprint> <input onclick="document.all.WebBrowser.ExecWB(8,1)" type="button" value="ҳ������">
&nbsp;<input onclick="document.all.WebBrowser.ExecWB(7,1)" type="button" value="��ӡԤ��">&nbsp;<input onclick="document.all.WebBrowser.ExecWB(6,1)" type="button" value="��ӡ">
 </center>
 <strong>������ģ������<%= i %> </strong>
</body>
</html>
<%
Call CloseRs(rs)
call closeConn()
%>
