<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
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
<form id="form1" name="form1" method="post" action="">
  <table width="95%" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="106">���������˲�ѯ</td>
      <td width="314"><label>
        <input name="TxtKey" type="text" id="TxtKey" />
        <input type="submit" name="Submit" value="�� ѯ" />
      </label></td>
	  <td align="right">	  	  <input type="button" name="Submit222" onclick="javascript:window.location.href='itevent.asp';" style="font-weight:bold; height:30px; width:120px;" value="�豸ά������ " /></td>
      <td width="227">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<table border="0" style="margin-left:30px;" cellpadding="4" cellspacing="1" bordercolor="#000000" width="95%" bgcolor="#F1F1F1" id="mytable">
  <tr>
    <td colspan="9" align="center" style=" font-weight:bold;">�豸ά��������ѯ&nbsp;</td>
  </tr>

  <tr align="center" style="font-weight:bold;">
    <td >���</td>
    <td >������</td>
    <td >����ʱ��</td>
    <td >��������</td>
	<td >������</td>
	<td >����״̬</td>	
    <td >���״̬</td>
	<% if session("Power")<999 then %>
	<td >����</td>
	<% end if %>
	<% if session("user_l")="11" then%>
	<td >������ϸ</td>
	<%
	end if
	%>
  </tr>

   <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=30
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
TxtKey=Trim(Request.Form("TxtKey"))
If TxtKey="" Then
	call CreateRs(rs,"select * from EPevent where ״̬='δ����' Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from EPevent Where comuser like '%"&TxtKey&"%' and ״̬='δ����' Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.Write("û�д˼�¼")
			response.End()
		end if
End If	
'ѭ����ȡ����
i=0
Do While Not rs.eof
call createRs(rsm,"select *from HUDSON_User where UserID='"& rs("comuser") &"'",1,1)
if not rsm.eof then
comuser=rsm("username")
else
comuser=""
end if
call closeRs(rsm)
%>  
  <tr align="center">
    <td ><%= i+1 %>&nbsp;</td>
    <td><%= comuser %>&nbsp;</td>
    <td><%= rs("����ʱ��") %>&nbsp;</td>
    <td><%= rs("��������") %>&nbsp;</td>
	<td><%= rs("jj")  %>&nbsp;</td>
    <td><%= rs("״̬")  %>&nbsp;</td>	
	<% if session("user_l")="19" then%>
	<td><%= rs("chk") %></td>
	<% if session("Power")<999 then %>
	<td><a href="chk_ok.asp?id=<%= rs("id") %>">���</a></td>
	<% end if %>
	<td><a href="queryIT.asp?id=<%= rs("id")%>">��ϸ�鿴</a></td>
	<% end if%>
  </tr>
    <%
	i=i+1
	Rs.MoveNext
Loop	
%>
</table>
<p>&nbsp;</p>
<%
Call CloseRs(rs)
call closeConn()
%>
</body>
</html>







