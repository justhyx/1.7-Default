<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ģ���絥��ѯ</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body>

<table align="right"><tr><td><a href="addtrymolde.asp?event=add">��ģ���絥���</a></td></tr></table>
          <form id="form1" name="form1" method="post" action="search.asp">
		  <p>&nbsp;</p>
           <table style="margin-left:100px"><tr><td>
			 <select name="stylename">
			<option value="ģ�ߵ���">ģ�ߵ���</option>\
			<option value="Ʒ��">��/Ʒ��</option>
			<option value="��׼Ҫ����̨">��׼Ҫ����̨</option>
			<option value="��ģ����">��ģ����</option>
			</select> &nbsp;
            <input name="TxtKey" type="text" id="TxtKey" class="noprint" />
            <input type="submit" name="Submit" value="�� ѯ" class="noprint" />
			</td></tr></table>
          </form>  
  <table id="mytable" width="1000" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center" style="font-weight:bold;">
	  <td>���</td>
      <td>��ģ����</td>
      <td>��/Ʒ��</td>
      <td>��׼Ҫ����̨</td>
      <td>����</td>
      <td>����</td>
      <td>ģ�ߵ���</td>
	  <td>��ϸ��Ϣ</td>
    </tr>
    <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=40
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
Dim TxtKey,stylename
stylename=Trim(request.form("stylename"))
TxtKey=Trim(Request.form("TxtKey"))
If TxtKey="" Then
		call CreateRs(rs,"select * from trymolde",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where "+stylename+" like '%"& TxtKey &"%' Order By addtime",1,1)
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
i=1+(ThisCurrentPage-1)*ThisPageSize
Do While Not Rs.EOF and i<ThisPageSize+1+(ThisCurrentPage-1)*ThisPageSize
%>
    <tr align="center">
	  <td><%=i%></td>
      <td><%= rs("��ģ����") %> | <%= rs("����") %></td>
      <td><%= rs("Ʒ��") %>&nbsp;</td>
      <td><%= rs("��׼Ҫ����̨") %>&nbsp;</td>     
      <td><%= rs("����") %>&nbsp;</td>
      <td><%= rs("addtime") %></td>
      <td><%= rs("ģ�ߵ���") %></td>
      <td><a href="addtrymolde.asp?numberid=<%=rs("id")%>&event=change">��ϸ��Ϣ</a></td>
    </tr>
    <%
		i=i+1
		Rs.MoveNext
		Loop	
	%>
    <tr>	
      <td colspan="15" align="center">&nbsp;</td>
    </tr>
</table>
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
</body>
</html>
<%
Call CloseRs(rs)
call closeConn()
%>