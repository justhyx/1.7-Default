<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ģ���絥����</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
Function getFormatDateString(dtNumber)
 Dim strDate
 If isNumeric(dtNumber) Then
  If dtNumber<10 Then
   dtNumber = "0" & dtNumber
  End If
  strDate = dtNumber
 Else
  strDate = "" 
 End if
 getFormatDateString = strDate
End Function
yy=year(now)
mm=getFormatDateString(month(now))
dd=getFormatDateString(int(day(now))+1)
dastr=yy&"-"&mm&"-"&dd
 %>
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
<% if session("UserName")="����" then  %>
<body>
          <form id="form1" name="form1" method="post" action="arrange.asp">
            <label>
            <input name="TxtKey" type="text" id="TxtKey" class="noprint" />
            </label>
            <label>
            <input type="submit" name="Submit" value="�� ѯ" class="noprint" />
            </label>
            <label class="noprint">�������ڲ�ѯ ����2012-06-07 ע����������� | <a href="arrange_a.asp">������ģ����</a></label>
          </form>  
  <table id="mytable" width="2000" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center" style="font-weight:bold;">
      <td>��ģ����</td>
      <td>��/Ʒ��</td>
      <td>��׼Ҫ����̨</td>
      <td>Ѩ��</td>
      <td>ȷ�ϻ�̨</td>
      <td>����ʱ��</td>
      <td>��ע</td>
      <td>����</td>
      <td>��������</td>
      <td>����</td>
      <td>��ע</td>
      <td>����</td>
      <td>���ϱ��</td>
      <td>ģ�ߵ���</td>
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
		call CreateRs(rs,"select * from trymolde where ����='"&dastr&"' Order By ����ȷ�ϻ�̨ ASC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where ���� like '%"&TxtKey&"%' Order By ����ȷ�ϻ�̨ ASC",1,1)
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
	<form id="form<%= i %>" name="form<%= i %>" method="post" action="chk.asp?action=arrange">
    <tr>
      <td><%= rs("��ģ����") %> | <%= rs("����") %></td>
      <td><%= rs("Ʒ��") %>&nbsp;</td>
      <td><%= rs("��׼Ҫ����̨") %>&nbsp;</td>
      <td><%= rs("Ѩ��") %>&nbsp;</td>
      <td><input name="okji" type="text" id="okji" value="<%= rs("����ȷ�ϻ�̨") %>" size="10" /></td>
      <td><input name="arrangetime" type="text" id="arrangetime" value="<%= rs("����ʱ��") %>" size="10" /></td>
      <td><input name="sgnote" type="text" id="sgnote" value="<%= rs("���ܱ�ע") %>" size="25" />
        <input name="id" type="hidden" id="id" value="<%= rs("id") %>" />
      <input type="submit" name="Submit2" value="�ύ" /></td>
      <td><%= rs("����") %>&nbsp;</td>
      <td><%= rs("��������") %></td>
      <td><%= rs("����") %></td>
      <td><%= rs("��ע") %></td>
      <td><%= rs("addtime") %></td>
      <td><label>
        <%= rs("���ϱ��") %></label></td>
      <td><%= rs("ģ�ߵ���") %></td>
      <td><label><%= rs("ģ��״̬") %></label></td>
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