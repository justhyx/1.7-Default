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
.STYLE1 {font-size: 24px}
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
<table width="500" border="0" cellpadding="4" cellspacing="1" style="margin-left:50px" >
  <tr>
    <td colspan="3">&nbsp; </td>
  </tr>
  <tr>
  <td><a href="computer02.asp" target="frmright"><span style="font-weight:bold;">���Թ����¼</span></a></td>
  <td><a href="computer03.asp" target="frmright"><span style="font-weight:bold;color:#FF0000;">����ʹ����ϸ</span></a></td>
  <td><a href="computer04.asp" target="frmright"><span style="font-weight:bold;">����ά����¼</span></a></td>
  </tr>
</table>
<table border="0" style="margin-left:50px;" cellpadding="4" cellspacing="1" bordercolor="#000000" width="100%" id="mytable" bgcolor="#F1F1F1">
<tr><td colspan="8" style="margin-left:50px;">
<form id="form2" name="form2" method="post" action="comchk.asp?event=search02">
      <select name="ziduan" id="ziduan">
      <option value="ʹ����">ʹ����</option>
      <option value="MAC">�����ַ</option>
	  <option value="ʹ�ò���">ʹ�ò���</option>
      <option value="������">������</option>
	  <option value="�칫��">�칫��</option>
	  <option value="ipdress">IP��ַ</option>
	  <option value="˫IP">��IP</option>
	  <option value="��ע">��ע</option>
    </select>
	&nbsp;<input type="text" id="checkin" name="checkin" size="32">
	&nbsp;<input type="submit" name="submit" value="�� ��">
	</form></td></tr>
	<%
comsql=request.cookies("comsql")
if comsql="" then
call createRS(rs,"select * from ����ʹ����ϸ where ����ʹ��ʱ�� is null",1,1)
else 
call createRS(rs,comsql,1,1)
end if
response.cookies("comsql")=""
response.cookies("comsql").Expires=dateadd("h",24,now())
%>
<tr>
<td>�����ַ</td>
<td>IP��ַ</td>
<td>˫IP</td>
<td>ʹ����</td>
<td>ʹ�ò���</td>
<td>�칫��</td>
<td>������</td>
<td>��ʼʹ��ʱ��</td>
<td>����ʹ��ʱ��</td>
<td>��ע</td>
<td>����</td>
</tr>
<%
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


ThisRsCount=Rs.RecordCount
'����ÿҳ��ʾ��������¼
Rs.PageSize=ThisPageSize
'��ȡ��ҳ��
ThisPageCount=Rs.PageCount
'�жϵ�ǰҳ�����Ч��
If ThisCurrentPage>ThisPageCount Then ThisCurrentPage=ThisPageCount 
'���õ�ǰҳ��
if ThisCurrentPage>0 then
Rs.AbsolutePage=ThisCurrentPage
end if
%>
<%
if not rs.eof then
i=1+(ThisCurrentPage-1)*ThisPageSize
Do While Not Rs.EOF and i<ThisPageSize+1+(ThisCurrentPage-1)*ThisPageSize
%>
<tr>
<td><%=rs("MAC")%></td>
<td><%=rs("ipdress")%></td>
<td><%=rs("˫IP")%></td>
<td><%=rs("ʹ����")%></td>
<td><%=rs("ʹ�ò���")%></td>
<td><%=rs("�칫��")%></td>
<td><%=rs("������")%></td>
<td><%=rs("��ʼʹ��ʱ��")%></td>
<td><%=rs("����ʹ��ʱ��")%></td>
<td><%=rs("��ע")%></td>
<td width="13%"><a href="changeone.asp?id=<%=rs("id")%>" target="frmright">�� ��</a>&nbsp;<a href="comchk.asp?event=del02&id=<%=rs("id")%>" target="frmright">ɾ ��</a></td>
</tr>
<%
rs.movenext
i=i+1
loop
else 
%>
<tr><td colspan="8">δ�ҵ�����Ҫ��ļ�¼����������������</td></tr>
<%
end if
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
</body>
</html>







