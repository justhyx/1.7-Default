<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form id="form1" name="form1" method="post" action="piclist.asp">
  <table width="95%" border="0" align="center" cellpadding="8" cellspacing="0">
    <tr>
      <td width="9%">������Ʒ����</td>
      <td width="91%"><label>
        <input name="TxtKey" type="text" id="TxtKey" />
        <input type="submit" name="Submit" value="�� ѯ" />
        </label></td>
    </tr>
  </table>
</form>
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1">
  <tr align="center">
    <td><strong>��Ʒ����</strong></td>
    <td><strong>��̨</strong></td>
    <td><strong>ע�ܳ�������</strong></td>
    <td><strong>��¼ʱ��</strong></td>
  </tr>
  <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=20
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
'�����ѯ�ؼ��ֱ�������URL����л�ȡ��ѯ�Ĺؼ���TxtKey������ϳ�ģ����ѯSQL���
Dim TxtKey,agree,ArrangeTime
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from zs_condition Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from zs_condition Where zx_number like '%"&TxtKey&"%' Order by id DESC",1,1)
End If
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
'ѭ����ȡ����
i=0
Do While Not Rs.EOF and i<ThisPageSize
%>
  <tr align="center">
    <td><%= rs("zx_number") %>&nbsp;</td>
    <td><%= rs("zx_jt")%>&nbsp;</td>
    <td><a href="<%= rs("zx_address")%>"  target="_blank"><%= rs("zx_number") %>��������</a>&nbsp;</td>
    <td><%= rs("uptime")%>&nbsp;</td> 
  </tr>
  <%
	i=i+1
	Rs.MoveNext	
Loop	
%>
</table>
<p>&nbsp;</p>
<table width="760" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>�ܹ�<%= ThisRsCount %>����¼,<%= ThisPageSize %>����¼/ҳ,�ܹ�<%= ThisPageCount %>ҳ,��ǰ��<%= ThisCurrentPage %>ҳ</td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      ��һҳ
      <% Else %>
      <a href="?Page=1&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      ��һҳ
      <% Else %>
      <a href="?Page=<%=ThisCurrentPage-1%>&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      ��һҳ
      <% Else %>
      <a href="?Page=<%=ThisCurrentPage+1%>&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      ���һҳ
      <% Else %>
      <a href="?Page=<%=ThisPageCount%>&TxtKey=<%=TxtKey%>" class="blue">���һҳ</a>
      <% End If %>
    </td>
  </tr>
</table>
<%
	Call CloseRs(rs)
	call closeConn()
 %>
</body>
</html>
