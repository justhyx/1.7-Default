<style type="text/css">
<!--
.STYLE1 {color: #00CCCC}
-->
</style>
<% 
Dim myArray(7),k
	k=0
call CreateRs(rsz,"select * from meetroom",1,1)
	do while not rsz.eof
 			if  rsz("flog")=2 then
				Myarray(k)="��Ԥ��"
			elseif  rsz("flog")=1 then
				Myarray(k)="ʹ����"
			elseif  rsz("flog")=0 then
				Myarray(k)="����"
			end if
	rsz.movenext
	k=k+1
	loop
	call closeRs(rsz)		
	%>
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr>
    <td colspan="5" align="center">���������һ����</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>����</td>
    <td>״̬</td>
    <td>λ��</td>
    <td>��ʩ</td>
    <td>��������</td>
  </tr>
  <tr>
    <td>1�Ż�����</td>
    <td><%= Myarray(0) %>&nbsp;</td>
    <td>1�ų���һ¥</td>
    <td rowspan="4"><p>����Ӱװ�����</p>
        <p>�ͶӰ�Ƕ�̨</p></td>
    <td>6��</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>2�Ż�����</td>
    <td><%= Myarray(1) %></td>
    <td>1�ų���ģ�߼в��¥</td>
    <td>8��</td>
  </tr>
  <tr>
    <td>3�Ż�����</td>
    <td><%= Myarray(2) %></td>
    <td>1�ų���ģ�߼в��¥</td>
    <td>12��</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>4�Ż�����</td>
    <td><%= Myarray(3) %></td>
    <td>1�ų���ģ�߼в��¥</td>
    <td>12��</td>
  </tr>
  <tr>
    <td>5�Ż�����</td>
    <td><%= Myarray(4) %></td>
    <td>1�ų�����¥</td>
    <td><p>����Ӱװ�һ��</p>
        <p>�̶�ͶӰ��һ̨</p></td>
    <td>14��</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>6�Ż�����</td>
    <td><%= Myarray(5) %></td>
    <td>2�ų������ΰ칫����¥</td>
    <td rowspan="2"><p>����Ӱװ�һ��</p>
        <p>�ͶӰ��һ̨</p></td>
    <td>8��</td>
  </tr>
  <tr>
    <td>7�Ż�����</td>
    <td><%= Myarray(6) %></td>
    <td>2�ų������ΰ칫����¥</td>
    <td>6��</td>
  </tr>
  <tr>
    <td>8�Ŷ๦����</td>
    <td><%= Myarray(7) %></td>
    <td>�๦����</td>
    <td>&nbsp;</td>
	<td>100��</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="5" align="center"><img src="../img/map.gif" width="400" height="200" /></td>
  </tr>
</table>
