<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
dim conn,str
	set conn=server.CreateObject("adodb.connection")
		str="provider=microsoft.Jet.OLEDB.4.0;data source="& server.MapPath("db/hudsonwwwroot.mdb")
		conn.open str
		
sub CreateRs(rs,sql,n1,n2)
	set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn,n1,n2
end sub

sub closeRs(rs)
	rs.close
	set rs=nothing
end sub

sub closeConn()
	conn.close
	set conn=nothing
end sub
 %>
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
</style>
<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="396" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
    <td colspan="4" align="center" bgcolor="#E5E2DA" style="font-weight:bold;">������ʹ��״��</td>
  </tr>
  <tr>
    <td bgcolor="#E5E2DA">������</td>
    <td bgcolor="#E5E2DA">ʹ��ʱ��</td>
    <td bgcolor="#E5E2DA">�����Һ�</td>
    <td bgcolor="#E5E2DA">״̬</td>
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
Dim TxtKey,agree,ArrangeTime
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from meeting Order By id DESC",1,1)
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
    <td bgcolor="#E5E2DA"><%= rs("M_meetingroom") %>&nbsp;</td>
    <td bgcolor="#E5E2DA"><%= rs("M_flog") %>&nbsp;</td>
  </tr>
        <%
	i=i+1
	Rs.MoveNext	
Loop	
%>
</table>
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>
