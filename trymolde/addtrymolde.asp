<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ģ���絥����</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
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
<%
ev=request.QueryString("event")
if ev="add" then
actvalue="addcentent"
 keypress="��  ��"
else
actvalue="exchange"
 keypress="��  ��"
end if


%>
</script>
<body>
<table width="1050" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
	<tr>
		<td align="left">
<form id="form1" name="form1" method="post" action="chk.asp?action=addnumber">
  <table width="950" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td height="50" colspan="4" align="center" bgcolor="#FFFFFF" style="font-size:24px; font-weight:bold; font-family:'����'">��ģ���絥</td>
    </tr>
    <tr>
      <td width="137" height="20" align="right" bgcolor="#FFFFFF" >��/Ʒ��
      <label></label></td>
      <td width="320" align="left" bgcolor="#FFFFFF" ><input name="t2" type="text" id="t2" />
      <input type="submit" name="Submit2" value="��ȡ��Ʒ��Ϣ" /></td>
      <td width="301" align="left" bgcolor="#FFFFFF" ><a href="search.asp">��ѯ��ģ���絥</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://192.168.1.7/trymolde/print.asp">��ģ���絥��ӡ</a></td>
      <td align="right" bgcolor="#FFFFFF" ><p>&nbsp;</p>        </td>
    </tr>
  </table> 
</form>
<form id="form2" name="form2" method="post" action="chk.asp?action=<%=actvalue%>">
     <%
	 dim at,bt,ct
	 'ndd=Request.QueryString("numberid")
	 ndd=Cint(Request.QueryString("numberid"))
	if request.QueryString("event")="add" then
	 call createRs(rsm,"select *from trymolde where id='"& ndd &"'",1,1)
	 if not rsm.eof then
	 numberidadd=rsm("Ʒ��")
	 end if
	 call closeRs(rsm)
	 if numberidadd="" then
	 	at=""
		bt=""
		cl=""
		dt=""
		et=""	
	else
		call CreateRs(rsa,"select * from compose where ��Ʒ����='"& numberidadd &"'",1,1)
			if rsa.eof then
				at=""
				bt=""
				cl=""
				dt=""
				et=numberidadd
				nq=""
				szsl=""
				yl=""
				ylf=""
				posl=""
				bz=""
			else	
				at=rsa("Ѩ��")
				bt=rsa("��Ʒ���ϱ��")
				cl=rsa("����")
				dt=rsa("��λ")
				et=rsa("��Ʒ����")
				nq=rsa("����")
				szsl=rsa("��������")
				yl=rsa("����")
				ylf=rsa("����F")
				posl=rsa("PO")
				bz=rsa("��ע")
				mjdd=rsa("ģ�ߵ���")
				smlx=rsa("��ģ����")
				lx=rsa("����")
				poc=rsa("PO����")
			end if
			call closeRs(rsa)
	end if
elseif request.QueryString("event")="change" then
			call CreateRs(rsa,"select * from trymolde where id='"& ndd &"'",1,1)
			if not rsa.eof then
				at=rsa("Ѩ��")
				bt=rsa("���ϱ��")
				cl=rsa("����")
				et=rsa("Ʒ��")
				nq=rsa("����")
				dt=rsa("��׼Ҫ����̨")
				szsl=rsa("��������")
				yl=rsa("����")
				ylf=rsa("����F")
				posl=rsa("PO")
				bz=rsa("��ע")
				mjdd=rsa("ģ�ߵ���")
				smlx=rsa("��ģ����")
				lx=rsa("����")
				poc=rsa("PO����")
				end if
	end if
 %>
  <table width="850" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td colspan="4" bgcolor="#FFFFFF"><hr width="100%" size="1" noshade="noshade" /></td>
    </tr> 
	</table>
  <table width="850" border="1" cellpadding="4" cellspacing="0" bordercolor="#030303" bgcolor="#CCCCCC">
    <tr>
      <td width="150" align="right" bgcolor="#FFFFFF">��/Ʒ��</td>
      <td width="300" colspan="2" bgcolor="#FFFFFF"><label></label>
      <input name="tx" type="text" id="tx" value="<%= et %>" disabled="true"/><input type="hidden" name="tryid" value="<%= ndd%>"></td>
      <td width="100" align="right" bgcolor="#FFFFFF">ģ�ߵ���</td>
      <td width="150" bgcolor="#FFFFFF"><label>
      <select name="t1" id="t1">
	<option value="<%=mjdd%>" selected="selected"><%=mjdd%></option>
        <option value="�ŷ�">�ŷ�</option>
        <option value="����">����</option>
        <option value="�߽�">�߽�</option>
	<option value="����">����</option>
	<option value="�޲���">�޲���</option>
	<option value="����">����</option>
	<option value="Ĳ���">Ĳ���</option>
        <option value="��־��">��־��</option>
        <option value="�Ž�">�Ž�</option>
        <option value="��Ӣ��">��Ӣ�� </option>
        <option value="Ҧ��">Ҧ��</option>
	<option value="κ���">κ���</option>
        <option value="��ѧ��">��ѧ��</option>
	<option value="������">������</option>
	<option value="�ձ���">�ձ���</option>
	<option value="������">������</option>
	<option value="�ⷢ��">�ⷢ��</option>
              </select>
      </label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">��ģ����</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="s1" id="s1">
		<option value="<%=smlx%>" selected="selected"><%=smlx%></option>
          <option value="�¹�">�¹���ģ</option>
          <option value="����">������ģ</option>
        </select>
      </label></td>
      <td bgcolor="#FFFFFF">ģ������
        <select name="mold_type" id="mold_type">
          <option value="����">����ģ��</option>
          <option value="����">����ģ��</option>
        </select>
        </td>
      <td align="right" bgcolor="#FFFFFF">��׼Ҫ����̨</td>
      <td bgcolor="#FFFFFF"><input name="t3" type="text" id="t3" value="<%= dt %>" /></td>
    </tr>
    <tr>
	  <td align="right" bgcolor="#FFFFFF">Ѩ��</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="t5" type="text" id="t5" value="<%= at %>" /></td>
      <td align="right" bgcolor="#FFFFFF">���ϱ��</td>
      <td bgcolor="#FFFFFF"><input name="t6" type="text" id="t6" value="<%= bt %>" /></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">����</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="t7" type="text" id="t7" value="<%=cl %>" size="60" /></td>
      <td align="right" bgcolor="#FFFFFF">ԭ��
        <input name="t82" type="text" id="t82" size="5" value="<%=yl%>"/>
        KG</td>
      <td align="left" bgcolor="#FFFFFF">������
        <input name="t822" type="text" id="t822" size="5" value="<%=ylf%>"/>
        KG</td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">����</td>
      <td colspan="2" bgcolor="#FFFFFF">
        <input name="t9" type="text" id="t9" value="<% =nq %>"onClick="WdatePicker()" style="hand();" />
        <img onclick="WdatePicker({el:$dp.$('t9')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle"></td>
      <td align="right" bgcolor="#FFFFFF">��������</td>
      <td bgcolor="#FFFFFF"><input name="t8" type="text" id="t8" value="<% =szsl%>"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">����</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="t11" id="t11">
	  <option value="<%=lx%>" selected="selected"><%=lx%></option>
        <option value="����">����</option>
        <option value="�ⷢ">�ⷢ</option>
        <option value="����">����</option>
      </select>
      �����豸
      <select name="fzsb" id="fzsb">
        <option value="ģ��">ģ��</option>
        <option value="������">������</option>
        <option value="ģ�� ������">ģ�� ������</option>
        <option value="0" selected="selected">��ѡ��</option>
      </select></td>
      <td align="right" bgcolor="#FFFFFF">PO����
        <input name="t12" type="text" id="t12" size="5" value="<%= posl%>"/></td>
      <td bgcolor="#FFFFFF">�ڿ��Ʒ����
        <select name="POchk" id="POchk">
		  <option value="" selected="selected"></option>
		  <option value="<%=poc%>" selected="selected"><%=poc%></option>
          <option value="����">����</option>
          <option value="����">����</option>
              </select></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">��ע</td>
      <td colspan="2" bgcolor="#FFFFFF"><textarea name="b1" rows="4" id="b1" value="<%=bz%>"></textarea></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="<%= keypress%>" />
        <label>
<input name="radiobutton" type="checkbox" id="radiobutton" value="radiobutton" checked="checked" />        
����</label>        </tr> 
</table>
</form>
</td>
</tr>
</table>	
  <hr align="center" width="1050" size="2" noshade="noshade" />
  <table width="1250" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#030303" bgcolor="#F1F1F1" id="mytable">
    <tr align="center" style="font-weight:bold;">
      <td>��ģ����</td>
      <td>��/Ʒ��</td>
      <td>Ҫ����̨</td>
	  <td>�����豸</td>
      <td>Ѩ��</td>
      <td>���ϱ��</td>
      <td>����</td>
      <td>��������</td>
      <td>ԭ��</td>
      <td>������</td>
      <td>PO</td>
	  <td>�ڿ��Ʒ����</td>
      <td>����</td>
      <td>ģ�ߵ���</td>
      <td>��ע</td>
      <td>����</td>
      <td>����</td>
    </tr>
<%
	Dim TxtKey
	TxtKey=Request("TxtKey")
	If TxtKey="" Then
		dress=Request.ServerVariables("REMOTE_ADDR")
		call CreateRs(rs,"select * from trymolde where okfolg='"&int(1)&"' and dress='"&dress&"' and addtime='"&date()&"' and (CONVERT(datetime, ����) >= DATEADD(day, - 1, GETDATE())) Order By id DESC ",1,1)
			if rs.eof or rs.bof then
				response.End()
			end if
	Else
		call CreateRs(rs,"select * from trymolde Where Ʒ�� like '%"&TxtKey&"%' and (CONVERT(datetime, ����) > GETDATE()) Order By ID DESC",1,1)
			if rs.eof or rs.bof then
				response.Write("û�д˼�¼")
				response.End()
			end if
	End If	
	i=0
	Do While Not Rs.EOF
	%>
    <tr>
      <td><%= rs("��ģ����") %> | <%= rs("����") %> | <%= rs("mold_type") %></td>
      <td><%= rs("Ʒ��") %>&nbsp;</td>
      <td><%= rs("��׼Ҫ����̨") %>&nbsp;</td>
	  <td><%= rs("fzsb") %>&nbsp;</td>
      <td><%= rs("Ѩ��") %>&nbsp;</td>
      <td><%= rs("���ϱ��") %>&nbsp;</td>
      <td><%= rs("����") %>&nbsp;</td>
      <td><%= rs("��������") %>&nbsp;</td>
      <td><%= rs("����") %>&nbsp;</td>
      <td><%= rs("����F") %>&nbsp;</td>
      <td><%= rs("PO") %></td>
	  <td><%= rs("PO����") %></td>
      <td>&nbsp;<%= rs("����") %></td>
      <td><%= rs("ģ�ߵ���") %>&nbsp;</td>
      <td><%= rs("��ע") %></td>
      <td><%= rs("addtime") %></td>
      <td><% if session("UserName")="����" or session("UserName")="���޾�" then  %>
          <a href="chk.asp?id=<%= rs("id")%>&amp;action=modify&op=isok&ipdress=addtrymolde">����</a> |
        <% end if %>
		 <a href="addtrymolde.asp?numberid=<%= rs("id")%>&amp;event=change" >����</a> |
        <a href="chk.asp?id=<%= rs("id")%>&amp;action=modify&op=isno&ipdress=addtrymolde" onclick="return confirm('ȷ��Ҫȡ��������¼��')">ȡ��</a></td>
    </tr>
    <%
	i=i+1
	Rs.MoveNext
	Loop	
	%>
</table>
</body>
</html>
<%
call closeConn()
%>



