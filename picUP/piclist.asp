<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
if session("UserName")="����" or session("UserName")="����" or session("UserName")="�߽�" or session("UserName")="�ŷ�" or session("UserName")="��ӢȨ" then 
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i)
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function
%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form id="form1" name="form1" method="post" action="">
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
    <td><strong>ͼ������</strong></td>
    <td><strong>�汾��</strong></td>
    <td><strong>��¼��</strong></td>
    <td><strong>��¼ʱ��</strong></td>
    <td><strong>������</strong></td>
    <td><strong>����ʱ��</strong></td>
    <td><strong>ͼ��״̬</strong></td>
    <td><strong>����</strong></td>
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
	call CreateRs(rs,"select * from PicFiles Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from PicFiles Where name like '%"&TxtKey&"%' Order by newfolg DESC",1,1)
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
    <td><%= rs("Name") %>&nbsp;</td>
    <td><%= rs("Suffix")%>&nbsp;</td>
    <td><%= rs("uppeople")%>&nbsp;</td>
    <td><%= rs("uptime")%>&nbsp;</td>
    <% if rs("okfolg")=1 then %>
    <td><a href="chkok.asp?id=<%= rs("id")%>&maseg=ok">����</a> | <a href="chkok.asp?id=<%= rs("id")%>&maseg=no" onclick="return confirm('��ȷ��ͼ�������𣿵��ȷ��������ɾ����ͼ������')">����</a>&nbsp;</td>
    <% elseif rs("okfolg")=0 then %>
    <td><%= rs("okpeople")%>&nbsp;</td>
    <% else %>
    <td>����&nbsp;</td>
    <% end if %>
    <td><%= rs("oktime")%>&nbsp;</td>
    <% if rs("newfolg")= 1 or rs("newfolg")= 2 then %>
    <td style="color:#FF0000;">��&nbsp;</td>
    <% elseif rs("newfolg")= 0 then %>
    <td>��&nbsp;</td>
    <% end if
	  	if rs("newfolg")= 1 then
			   %>
    <td><a href="<%= rs("DownAddress")%>">�������</a>&nbsp;</td>
    <% elseif rs("newfolg")= 0 then %>
    <td style="color:#999999;">�������</a>&nbsp;</td>
    <% end if %>
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

	else             '���ηǷ��û�
	response.Write"<script> alert('ʹ���û���׼ȷ���뷵��'); history.back(); </script>"
	response.End()
	end if
 %>
</body>
</html>
