<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
if session("UserName")="" then
	response.Redirect("../login.asp")
end if
 %>
<!--#include file="../connet/conn.asp" -->
<% call CreateRs(rs,"select * from UseCar where id="&Trim(Request.QueryString("id")),1,3) %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #3399FF}
.STYLE2 {color: #FF0000}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="post" action="check.asp?action=modify&id=<%= rs("id") %>">
  <table width="100%" border="0" align="center" cellpadding="8" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center">
      <td colspan="5" style="font-weight:bold">�ó������
	  <%  if rs("Urgencyfloy")=2 then %>
	  <span class="STYLE2">��</span>
	  <% end if %>	  </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="131" align="right">���벿�ţ�</td>
      <td width="219"><span class="STYLE1">
        <label><%= rs("CarDepartment") %></label>
      </span></td>
      <td width="103" align="right">��д�������ڣ� </td>
      <td colspan="2"><span class="STYLE1"><%= rs("addtime") %>&nbsp;</span></td>
    </tr>
    <tr>
      <td align="right">���복��Ԥ����ֹʱ�䣺</td>
      <td><span class="STYLE1"><%= rs("StartTime") %>&nbsp;</span></td>
      <td align="right">����</td>
      <td colspan="2"><span class="STYLE1"><%= rs("OverTime")%>&nbsp;</span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">Ŀ�ĵأ�</td>
      <td><span class="STYLE1"><%= rs("Destination")%>&nbsp;</span></td>
      <td align="right">Ԥ�ư�������ʱ�䣺</td>
      <td colspan="2"><span class="STYLE1"><%= rs("UseTime") %>&nbsp;��Сʱ</span></td>
    </tr>
    <tr>
      <td align="right">Ԥ�Ƴ���ʱ�䣺</td>
      <td><%= rs("AboutTime") %>&nbsp;</td>
      <td align="right">���복��</td>
      <td><%= rs("cartype") %>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	    <tr>
      <td align="right">�����ó�����</td>
      <td><%= rs("peopleqty") %>&nbsp;</td>
      <td align="right">Ԥ�ƿ�����</td>
      <td><%= rs("pl") %>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">�� �ɣ�</td>
      <td colspan="4"><span class="STYLE1"><%= rs("Reason") %></span></td>
    </tr>
    <tr>
      <td align="right">* ��ʱ�����ó����ɣ� </td>
      <td colspan="4"><span class="STYLE1"><%= rs("Urgency")  %>&nbsp;</span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">* ���ƴ�ʩ�� </td>
      <td colspan="4">
        <textarea name="Improve" rows="4" id="Improve"><%= rs("Improve") %></textarea>
      </td>
    </tr>
    <% if session("Power")<9999 then %>
    <tr>
      <td rowspan="2" align="right">�ó�����ǩ��</td>
      <td>*���³�/�ܾ���/���ܾ�����ʱ�����ó�ǩ�� </td>
      <td>���ž���</td>
      <td width="81">������</td>
      <td width="80">�����</td>
    </tr>
    <tr>
      <td>
	  <% if session("Power")<9 then %>
	  <label>
        <select name="UrgencyAgree" id="UrgencyAgree">
          <option value="1" selected="selected">ͬ��</option>
          <option value="0">��ͬ��</option>
        </select>
      </label>
	   <% end if %>	  </td>
      <td><select name="Agree" id="Agree">
          <option value="1" selected="selected">ͬ��</option>
          <option value="0">��ͬ��</option>
      </select></td>
      <td><%= rs("UsePeople") %>&nbsp;</td>
      <td><%= rs("Agreepeople") %>&nbsp;</td>
    </tr>  
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><label>
        <input type="submit" name="Submit" value="ȷ�ϳ���" />
      </label></td>
    </tr>
  </table>
    <% end if %>
</form>
  <% 
  if Trim(Request.QueryString("action"))="arrange" then
   %>
 <form id="form2" name="form2" method="post" action="check.asp?action=arrange&id=<%= rs("id") %>"> 
  <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td width="77" align="right">�������ƺţ�</td>
      <td width="604"><label>
        <select name="CarNumber" id="CarNumber">
          <option value="A691T" selected="selected">���� A691T</option>
		  <option value="6371S">��� 6371S</option>
          <option value="C3D46">Ƥ�� C3D46</option>
          <option value="G153Q">��ʮ�� G153Q</option>
		  <option value="B35J65">���� B35J65</option>
		  <option value=" ">���� </option>
		  
        </select>
      </label></td>
    </tr>
    <tr>
      <td align="right">˾����</td>
      <td><label>
      <select name="CarPeople" id="CarPeople">
        <option value="��˾�� 13682421190  65190">��˾��</option>
        <option value="��˾��  661311 13510333201">��˾��</option>
        <option value="�� 678815">��˾��</option>
		<option value=" ">����</option>
      </select>
      </label></td>
    </tr>
    <tr>
      <td align="right">���복��</td>
      <td><label>
      <input name="note" type="text" id="note" size="80" />
      </label></td>
    </tr>
    <tr>
      <td align="right"><label>
        <input type="submit" name="Submit2" value="ȷ�ϳ���" />
      </label></td>
      <td><label>
<input name="delfolg" type="radio" value="1" checked="checked" />
ȷ������
<input type="radio" name="delfolg" value="2" />
      ����</label></td>
    </tr>
  </table>
</form>
<% 
end if
call closeRs(rs)
 %>
 <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center" bgcolor="#FFFFFF">
    <td>���״̬</td>
	<td>�����</td>
    <td>�ɳ�״̬</td>
    <td>���ƺ�</td>
    <td>ʦ��/�绰</td>
   </tr>
<% 
call CreateRs(rs,"select * from UseCar where id="&Trim(Request.QueryString("id")),1,1)
 %>
  <tr align="center">
    <% if rs("Agree")="0" or rs("UrgencyAgree")="0" then %>
	<td><span class="STYLE1">���δͨ��</span></td>	
	<% elseif rs("Agree")="1" or rs("UrgencyAgree")="1" then %>
	<td><span class="STYLE1">ͬ��</span></td>	
	<% elseif rs("Urgencyfloy")>0 then  %>
	<td><span class="STYLE2 STYLE1">�ȴ����</span></td>
	<% end if %>
	<td><span class="STYLE2 STYLE1"><%= rs("Agreepeople") %></span></td>
	<% if rs("arrange")="1" then %>
	<td><span class="STYLE1">�ɳ��ȴ�</span></td>
	<% else %>
	<td><span class="STYLE1">�������</span></td>
	<% end if %>
    <td><span class="STYLE1"><%= rs("CarNumber") %>&nbsp;</span></td>
    <td><span class="STYLE1"><%= rs("CarPeople") %>&nbsp;</span></td>
   </tr>
</table>

 <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1">
   <tr>
     <td width="121" align="center">���복��</td>
     <td width="560"><%= rs("note") %>&nbsp;</td>
   </tr>
</table>
</body>
</html>
<% 
call closeRs(rs)
call closeConn()
 %>