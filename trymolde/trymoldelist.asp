<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ģ���絥����</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<script language="javascript">
window.onload=function showtable(){
   var tablename=document.getElementById("mytable");
   var li=tablename.getElementsByTagName("tr");
   for (var i=0;i<=li.length;i++){
    if(i%2==0){
     li[i].style.backgroundColor="#F1F1F1";
     li[i].onmouseout=function(){
         this.style.backgroundColor="#F1F1F1";
      }
    }else{
     li[i].style.backgroundColor="#FFFFFF";
     li[i].onmouseout=function(){
         this.style.backgroundColor="#FFFFFF"
      }
    }
      li[i].onmouseover=function(){
      this.style.backgroundColor="#00CCFF";
      }
      
   }
}
</script>
<body>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50" align="center" style="font-size:24px; font-family:'����'; font-weight:bold;">��ģ���絥</td>
  </tr>
  <tr>
    <td height="20" align="right" style="font-weight:bold;"><a href="print.asp">Ԥ��</a></td>
  </tr>
</table>
<table id="mytable" width="1050" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center" style="font-weight:bold;">
    <td>��ģ����</td>
    <td>��/Ʒ��</td>
    <td>����̨</td>
    <td>Ѩ��</td>
    <td>���ϱ��</td>
    <td>����</td>
    <td>��������</td>
    <td>����</td>
    <td>ģ�ߵ���</td>
    <td>��ע</td>
    <td>����</td>
	<td>����</td>
  </tr>
   <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=16
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
	call CreateRs(rs,"select * from trymolde where addtime='"&date()&"' and okfolg='"&int(1)&"' Order By id DESC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where Ʒ�� like '%"&TxtKey&"%' Order By ID DESC",1,1)
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
    <td><%= rs("��ģ����") %>&nbsp;</td>
    <td><%= rs("Ʒ��") %>&nbsp;</td>
    <td><%= rs("��׼Ҫ����̨") %>&nbsp;</td>
    <td><%= rs("Ѩ��") %>&nbsp;</td>
    <td><%= rs("���ϱ��") %>&nbsp;</td>
    <td><%= rs("����") %>&nbsp;</td>
    <td><%= rs("��������") %>&nbsp;</td>
    <td><%= rs("����") %>&nbsp;</td>
    <td><%= rs("ģ�ߵ���") %>&nbsp;</td>
    <td><%= rs("��ע") %></td>
    <td><%= rs("addtime") %></td>
	<td><% if session("UserName")="����" or session("UserName")="���޾�" then  %><a href="chk.asp?id=<%= rs("id")%>&action=modify&op=isok&ipdress=trymoldelist">����</a> | <a href="chk.asp?id=<%= rs("id")%>&action=modify&op=isno&ipdress=trymoldelist" onclick="return confirm('ȷ��Ҫȡ��������¼��')">ȡ��</a><% end if %></td>	
  </tr>
      <%
	i=i+1
	Rs.MoveNext
Loop	
%>
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
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>
