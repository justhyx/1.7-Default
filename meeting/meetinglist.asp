<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-color: #E5E2DA;
}
-->
</style></head>

<body>
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
    <td colspan="6" align="center" bgcolor="#E5E2DA" style="font-weight:bold;">������ʹ��״��</td>
  </tr>
  <tr>
    <td bgcolor="#E5E2DA">������</td>
    <td bgcolor="#E5E2DA">ʹ��ʱ��</td>
	<td bgcolor="#E5E2DA">���������</td>
    <td bgcolor="#E5E2DA">�ͻ�����</td>
    <td bgcolor="#E5E2DA">�����Һ�</td>
    <td bgcolor="#E5E2DA">״̬</td>
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
%>
<%
'�����ѯ�ؼ��ֱ�������URL����л�ȡ��ѯ�Ĺؼ���TxtKey������ϳ�ģ����ѯSQL���
Dim TxtKey,agree,ArrangeTime
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from meeting Order By ID DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from meeting Where M_usename like '%"&TxtKey&"%' Order By ID DESC",1,1)
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
    <td bgcolor="#E5E2DA"><%= rs("M_usename") %>&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_usetime") %>&nbsp;&nbsp;<%= rs("M_usetime_h") %>&nbsp;&nbsp;ʱ������<%= rs("M_overtime") %>&nbsp;&nbsp;<%= rs("M_overtime_h") %>&nbsp;&nbsp;ʱ&nbsp;</td>
	<td bgcolor="#E5E2DA"><%= rs("intention") %>&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_guest") %>&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_meetingroom") %></td>
    <td bgcolor="#E5E2DA"><a href="meetinglook.asp?id=<%= rs("id") %>"><%= rs("M_flog") %></a></td>
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
    <td align="center">
	<% If ThisCurrentPage=1 Then %>
		��һҳ
	<% Else %>
		<a href="?Page=1&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
	<% End If %>
	</td>
    <td align="center">
	<% If ThisCurrentPage=1 Then %>
		��һҳ
	<% Else %>
		<a href="?Page=<%=ThisCurrentPage-1%>&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
	<% End If %>
	</td>
    <td align="center">
	<% If ThisCurrentPage=ThisPageCount Then %>
		��һҳ
	<% Else %>
		<a href="?Page=<%=ThisCurrentPage+1%>&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
	<% End If %>
	</td>
    <td align="center">
	<% If ThisCurrentPage=ThisPageCount Then %>
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
