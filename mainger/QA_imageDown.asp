<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <p>&nbsp;</p>
  <table width="750" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="69">�ؼ��ֲ�ѯ</td>
      <td width="155"><label>
        <input name="TxtKey" type="text" id="TxtKey" />
      </label></td>
      <td width="498"><label>
        <input type="submit" name="Submit" value="�ύ" />
      </label></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<table width="760" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#EEEEEE">
  <tr>
    <td colspan="4" align="center" bgcolor="#86C4C4">��Ʒ�鿴</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">��Ʒ���</td>
    <td bgcolor="#FFFFFF">��Ʒ�鿴</td>
    <td bgcolor="#FFFFFF">�༭��</td>
    <td bgcolor="#FFFFFF">����ʱ��</td>
  </tr>
  <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=10
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
	call CreateRs(rs,"select * from QA_Check Order By ID DESC",1,1)
	if rs.eof or rs.bof then
		response.End()
	end if
Else
	call CreateRs(rs,"select * from QA_Check Where QA_ID like '%"&TxtKey&"%' Order By ID DESC",1,1)
	if rs.eof or rs.bof then
		response.Write"<script> alert('û�д˲�Ʒ������ϵ����'); history.back(); </script>"
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
    <td bgcolor="#FFFFFF"><%= rs("QA_ID") %><% if session("UserName")="����" or session("UserName")="�ź�ƻ" then %><a href="../up/chk.asp?id=<%= rs("QA_id")%>&objct=big_list">ɾ��</a><% end if %></td>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="4">
      <% 
	  call CreateRs(rsc,"select * from QA_image where QA_id='"&rs("QA_ID")&"'",1,1) 
	  DO while not rsc.eof
	  %>
	  <tr>
        <td><a href="../<%= rsc("QA_idimage") %>"><%= rsc("QA_name") %></a>&nbsp;</td>
      </tr>
      <% 
	  rsc.movenext
	  loop
	  call closeRs(rsc)
	   %>
    </table></td>
    <td bgcolor="#FFFFFF"><%= rs("QA_people") %>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%= rs("QA_addtime") %></td>
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
