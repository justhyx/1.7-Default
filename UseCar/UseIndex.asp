<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../Images/Admin.css" rel="stylesheet" type="text/css" />
 <script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
 <link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="check.asp?action=add">
  <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center">
      <td colspan="4" style="font-weight:bold">�ó������</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="129" align="right">���벿��</td>
      <td width="230"><label>
        <input name="CarDepartment" type="text" id="CarDepartment" value="<%= session("Department")%>" />
      </label></td>
      <td width="108" align="right">�����ˣ� </td>
      <td width="196"><label>
        <input name="UsePeople" type="text" id="UsePeople" value="<%= session("UserName")%>" />
      </label>        
        &nbsp;*����</td>
    </tr>
    <tr>
      <td align="right">Ԥ����ʼ����</td>
      <td><input name="StartTime" type="text" id="StartTime" onClick="WdatePicker()" style="hand();" />
        <img onclick="WdatePicker({el:$dp.$('StartTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle"></td>
      <td align="right">��</td>
      <td><input name="OverTime" type="text" id="OverTime" onClick="WdatePicker()" />
      <img onclick="WdatePicker({el:$dp.$('OverTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">Ŀ�ĵ�</td>
      <td><input name="Destination" type="text" id="Destination" /></td>
      <td align="right">Ԥ���ó�ʱ��</td>
      <td><input name="UseTime" type="text" id="UseTime" size="8" />
      Сʱ</td>
    </tr>
    <tr >
      <td align="right">Ԥ�Ƴ���ʱ��</td>
      <td><label>
<select name="hh" id="hh">
<option value=" ">��ѡ��</option>
  <% for i=0 to 24 %>
  <option value="<%= i %>"><%= i %></option>
  <% next %>
</select>      
ʱ
<select name="mm" id="mm">
  <option value="00:00">00</option>
  <option value="15:00">15</option>
  <option value="30:00">30</option>
  <option value="45:00">45</option>
</select>
      �֡�ע��δ��ǰ3Сʱ�����ó�����Ϊ�����ó���Ҫ��д��������</label></td>
      <td align="right">ѡ����</td>
      <td><label>
      <select name="cartype" id="cartype">
        <option value="�γ�">�γ�</option>
        <option value="����">����</option>
		<option value="Ƥ��">Ƥ��</option>
        <option value="1.5�ֻ���">1.5�ֻ���</option>
        <option value="5�ֻ���_����">5�ֻ���_����</option>
        <option value="5�ֻ���_������">5�ֻ���_������</option>
        <option value="9.8�ֻ���">9.8�ֻ���</option>
        <option value="0" selected="selected">��ѡ��</option>
      </select>
      </label></td>
    </tr>
    <tr >
      <td align="right">������</td>
      <td><label>
      <input name="peopleqty" type="text" id="peopleqty" />
      ��</label></td>
      <td align="right">Ԥ�Ƴ���������</td>
      <td><label>
        <input name="pl" type="text" id="pl" size="8" />
        ����
      </label></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">�� ��</td>
      <td colspan="3"><label>
        <textarea name="Reason" cols="40" rows="4" id="Reason"></textarea>
      *������д</label></td>
    </tr>
    <tr>
      <td align="right">* ��ʱ�����ó����� </td>
      <td colspan="3"><textarea name="Urgency" cols="40" rows="4" id="Urgency"></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">* ���ƴ�ʩ </td>
      <td colspan="3"><textarea name="Improve" cols="40" rows="4" id="Improve"></textarea></td>
    </tr>
	<% if session("Power")<9 then %>
    <tr>
      <td rowspan="2" align="right">�ó�����ǩ��</td>
      <td>*���³�/�ܾ���/���ܾ�����ʱ�����ó�ǩ�� </td>
      <td colspan="2">���ž���</td>
    </tr>
    <tr>
      <td>
	  <% if session("Power")<5 then %>
	  <label>
        <select name="UrgencyAgree" id="UrgencyAgree">
          <option value="1" selected="selected">ͬ��</option>
          <option value="0">��ͬ��</option>
        </select>
      </label>
	  <% end if %>	  </td>
      <td colspan="2"><select name="Agree" id="Agree">
        <option value="1" selected="selected">ͬ��</option>
        <option value="0">��ͬ��</option>
      </select></td>
    </tr>
	<% end if %>
  </table>
  <table width="100%" border="0" align="center" cellpadding="12" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>��ע��1���������ó���������ǰ�ƻ�����,����ǰ��д���ó����뵥��������ǩ��ú�ݽ��� <br />
���񲿣����ɰ��ŵ��ﳵ���� <br />
��ǰ����ʱ�����ޣ�a) ���������ó�������ǰһ���°�(��PM5:30)ǰ�ݽ�. <br />
b) ���������ó�,����ǰ3Сʱǰ�ݽ�. <br />
c) ���������ó�,���ڵ�������3:00ǰ�ݽ�. <br />
2�����������ó�,�ɲ��ž���ǩ����Ч. <br />
3: ��ʱ�����ó����ɲ��ž����(���³�/�ܾ���/���ܾ���)����һ�˹�ͬǩ����Ч. <br />
��ʱ�����ó����ϡ� * �� ��������д����Ч. <br />
4���������ռ��ڼ����ó���賵,������������PM5:30ǰ�ݽ����ó����뵥��,�������� <br />
���ž����(���³�/�ܾ���/���ܾ���)����һ�˹�ͬǩ����Ч. </td>
    </tr>
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr align="center">
      <td><label>
        <input name="action" type="hidden" id="action" value="add" />
        <input type="submit" name="Submit" value="�� ��" />
      </label></td>
    </tr>
  </table>
  <br />
</form>
</body>
</html>
