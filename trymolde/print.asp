<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/print.css" rel="stylesheet" type="text/css" />
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT>
</OBJECT>
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
dastr=year(now)&"-"&getFormatDateString(month(now))&"-"&getFormatDateString(day(now))
 %>
</head>
 <style media="print">
    .noprint { display: none }
  </style>
<body style="font-size:10px;">
<table width="1050" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000" id="mytable">
  <tr style="font-weight:bold;">    
    <td height="60" colspan="16"><table width="1050" height="60" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="4" align="center" style="font-size:24px; font-weight:bold; font-family:'����'">��ģ���絥</td>
      </tr>
      <tr>
        <td align="left"><%= date() %>&nbsp;
          <form id="form1" name="form1" method="post" action="">
            <label>
            ����������
            <input name="TxtKey" type="text" class="noprint" id="TxtKey" size="12" />
            </label>
            <label></label>
            <label class="noprint"> ����2012-06-07 �������̨</label>
            <label>
            <input name="equip_name" type="text" id="equip_name" size="12" />
            </label>
            <label class="noprint"> </label>
            <input type="submit" name="Submit" value="�� ѯ" class="noprint" />
          </form>          </td>
        <td width="200" align="left">���ϣ�</td>
      </tr>
    </table></td>
  </tr>
  <tr style="font-weight:bold;"> 
  <td>ID</td>   
    <td>����</td>
	<td>��/Ʒ��</td>
    <td>��̨</td>
	<td>�����豸</td>  
    <td>Ѩ��</td>
    <td>���ϱ��</td>
    <td>����</td>
	<td>ԭ��</td>
	<td>����</td>
    <td>����</td>
    <td>PO</td>
	<td>��Ʒ����</td>
    <td>����</td>
    <td>����</td>
    <td width="200">��ע</td>
  </tr>
  <%
'�����ѯ�ؼ��ֱ�������URL����л�ȡ��ѯ�Ĺؼ���TxtKey������ϳ�ģ����ѯSQL���
Dim TxtKey
TxtKey=Request("TxtKey")
equip=request.Form("equip_name")
If TxtKey="" and equip="" Then
sql="select * from trymolde Where ���� ='"&dastr&"'"
else
sql="select * from trymolde Where 1=1"
if TxtKey<>"" then
sql=sql&" and ���� like '%"&TxtKey&"%'"
end if
if equip<>"" then
sql=sql&" and ��׼Ҫ����̨='"& equip &"'"
end if
End If
sql=sql&" Order By ģ�ߵ��� DESC"
Call CreateRs(rs,sql,1,1)
		if rs.eof or rs.bof then
			response.write("û������")
			response.End()
		end if	
i=1
Do While Not Rs.EOF
%>
  <tr>
  	<td><%= i %></td>
    <td><%= rs("��ģ����") %> | <%= rs("����") %>| <%= rs("mold_type") %></td>
	<td width="60"><%= rs("Ʒ��") %></td>
    <td width="60"><%= rs("��׼Ҫ����̨") %></td>
	 <td><%= rs("fzsb") %></td>     
    <td width="40"><%= rs("Ѩ��") %></td>
    <td width="80"><%= rs("���ϱ��") %></td>
    <td width="80"><%= rs("����") %></td>
	<td><%= rs("����") %>KG</td>
	<td width="60"><%= rs("����F") %>KG</td>
    <td width="40"><%= rs("��������") %></td>
    <td><%= rs("PO") %></td>
	<td><%= rs("PO����") %></td>
    <td><%= right( rs("����"),5) %></td>
    <td><%= rs("ģ�ߵ���") %></td>
    <td><%= rs("��ע") %></td>
  </tr>
  <%  
	i=i+1
	Rs.MoveNext
	Loop	
	%>
</table>
   <center class=noprint> <input onclick="document.all.WebBrowser.ExecWB(8,1)" type="button" value="ҳ������">
&nbsp;<input onclick="document.all.WebBrowser.ExecWB(7,1)" type="button" value="��ӡԤ��">&nbsp;<input onclick="document.all.WebBrowser.ExecWB(6,1)" type="button" value="��ӡ">
 </center>
<strong>������ģ������<%= i - 1  %></strong>
</body>
</html>
<%
Call CloseRs(rs)
call closeConn()
%>



