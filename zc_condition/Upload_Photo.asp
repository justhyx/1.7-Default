<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<SCRIPT language=javascript>
<!--#include file="../connet/conn.asp" -->
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("��ѡ��Ҫ�ϴ����ļ�");
		document.form1.FileName.focus();
    	return false;
  	}
}
</SCRIPT>
</head>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-left: 0px;
}
-->
</style>
<%
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
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css">
<body>
<form action="Upfile_Photo.asp" method="post" name="form1" onSubmit="return check()" enctype="multipart/form-data">
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
          <tr>
            <td colspan="3" align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">ͼ���ϴ�</td>
          </tr>
          <tr>
            <td width="20%" align="right">��Ʒ���ţ�</td>
            <td colspan="2" align="left"><label>
              <input name="QA_ID" type="text" id="QA_ID" />
            </label></td>
          </tr>
          <tr>
            <td align="right">��̨��</td>
            <td colspan="2" align="left"><label>
              <input name="QA_name" type="text" id="QA_name" />
            </label></td>
          </tr>
          <tr>
            <td align="right">ͼ���ϴ���</td>
            <td colspan="2" align="left"><input name="FileName" height="30" type="file" size="30" /></td>
          </tr>
          <tr>
            <td align="right">&nbsp;</td>
            <td colspan="2" align="left"><label></label></td>
          </tr>
          <tr>
            <td align="right">��ע��</td>
            <td width="35%" align="left"><label>
              <textarea name="describe" rows="4" id="describe"></textarea>
            </label></td>
            <td width="45%" align="left" valign="bottom"><input type="submit" name="Submit" value="�� ��" /></td>
          </tr>
  </table>

</form>
<table id="mytable" width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center">
    <td><strong>��Ʒ����</strong></td>
    <td ><strong>��̨</strong></td>
    <td ><strong>��������</strong></td>
    <td ><strong>��¼ʱ��</strong></td>
    <td ><strong>����</strong></td>
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
	call CreateRs(rs,"select * from zs_condition Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from PicFiles Where name like '%"&TxtKey&"%' and okfolg='"&int(0)&"' Order by newfolg DESC",1,1)
			if rs.eof or rs.bof then
			response.Write("û�д˲�Ʒ")
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
  <tr align="center">
    <td><%= rs("zx_number") %>&nbsp;</td>
    <td><%= rs("zx_jt")%>&nbsp;</td>
    <td><a href="<%= rs("zx_address")%>"><%= rs("zx_number") %>��������</a>&nbsp;</td>
    <td><%= rs("uptime")%>&nbsp;</td>
    <td ><a href="chkok.asp?id=<%= rs("id")%>&amp;maseg=del" onclick="return confirm('��ȷ��ͼ�������𣿵��ȷ��������ɾ����ͼ������')">ɾ��</a></td>
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
    <td align="center"><% If ThisCurrentPage=1 Then %>
      ��һҳ
      <% Else %>
        <a href="?Page=1&amp;TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
        <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      ��һҳ
      <% Else %>
        <a href="?Page=<%=ThisCurrentPage-1%>&amp;TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
        <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      ��һҳ
      <% Else %>
        <a href="?Page=<%=ThisCurrentPage+1%>&amp;TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
        <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      ���һҳ
      <% Else %>
        <a href="?Page=<%=ThisPageCount%>&amp;TxtKey=<%=TxtKey%>" class="blue">���һҳ</a>
        <% End If %>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</thml>