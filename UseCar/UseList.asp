<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
session("Url")="../Usecar/UseList.asp"
if session("UserName")="" then
	response.Redirect("../login.asp")
end if
 %>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE2 {color: #FF0000}
-->
</style>
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
</head>

<body>
<a href="personnel_list.asp">�������</a>
<table id="mytable" width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center" style="font-weight:bold;">
    <td>���벿��</td>
    <td>������</td>
    <td><a href="uselist_addtime.asp">����ʱ��</a></td>
    <td><a href="UseList.asp?op=StartTime">�ó�ʱ��</a></td>
    <td>Ŀ�ĵ�</td>
    <td>��˽��</td>
    <td>���ʱ��</td>
    <td>ȷ�ϳ���</td>
    <td>����ʱ��</td>
    <td>�鿴</td>
  </tr>
   <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=21
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
	call CreateRs(rs,"select * from UseCar where delfolg<>'"&int(2)&"' and over_l='"& session("over_l") &"' Order BY ID DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from UseMacAddress Where UseName like '%"&TxtKey&"%'",1,1)
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
  <tr align="center" bgcolor="#FFFFFF">
    <td><%= rs("CarDepartment") %></td>
    <td><%= rs("UsePeople") %>&nbsp;</td>
    <td><%= rs("addtime") %></td>
    <td><%= rs("StartTime") %></td>
    <td><%= rs("Destination")%></td>
    <% if rs("Agree")="0" or rs("UrgencyAgree")="0" then %>
	<td>���δͨ��</td>	
	<% elseif rs("Agree")="1" or rs("UrgencyAgree")="1" then %>
	<td>ͬ��</td>	
	<% elseif rs("Urgencyfloy")>0 then  %>
	<td><span class="STYLE2">�ȴ����</span></td>
	<% end if %>
	<td><%= rs("ArrangeTime") %>&nbsp;</td>
	<% if rs("arrange")="1" then %>
	<td>
	<% if session("UserName")="ǮӰ÷" then %>
	<a href="UseLook.asp?action=arrange&id=<%= rs("id") %>"><font color="#FF0000">�ɳ��ȴ�</font></a>
	<% else %>
	<font color="#FF0000">�ɳ��ȴ�</font>
	<% end if %>
	</td>
	<% else %>
	<td>�������</td>
	<% end if %>
	<td><%= rs("addarrangetime") %>&nbsp;</td>
	<td><a href="UseLook.asp?id=<%= rs("id")%>">�鿴</a></td>
    <% if rs("Urgency")<>"" then %>
	<td><span class="STYLE2">��&nbsp;</span></td>
	<% end if %>
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
