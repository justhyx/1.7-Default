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
<% if session("UserName")="����¶" or session("UserName")="��սӮ" then  %>
<body>
  <table id="mytable" width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr style="font-weight:bold;">
      <td>��ģ����</td>
      <td>��/Ʒ��</td>
      <td>��������</td>
      <td>����</td>
      <td>ģ�ߵ���</td>
      <td>��ע</td>
      <td>����</td>
      <td>ģ��״̬</td>
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
Dim TxtKey
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from trymolde where okfolg='"&int(0)&"' and addtime='"&date()&"' Order By id DESC ",1,1)
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
	<form id="form<%= i %>" name="form<%= i %>" method="post" action="chk.asp?action=ments">
    <tr>
      <td><%= rs("��ģ����") %>&nbsp;</td>
      <td><%= rs("Ʒ��") %>&nbsp;</td>
      <td><%= rs("��������") %>&nbsp;</td>
      <td><%= rs("����") %>&nbsp;</td>
      <td><%= rs("ģ�ߵ���") %>&nbsp;</td>
      <td><%= rs("��ע") %></td>
      <td><%= rs("addtime") %></td>
      <td><label>
        <input name="id" type="hidden" id="id" value="<%= rs("id") %>" />
        <input name="moldements" type="text" id="moldements" value="<%= rs("ģ��״̬") %>" />
        <input type="submit" name="Submit2" value="�ύ" />
      </label></td>
    </tr>
	</form>
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
end if
%>
