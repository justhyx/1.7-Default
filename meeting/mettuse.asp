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
				Myarray(k)="已预订"
			elseif  rsz("flog")=1 then
				Myarray(k)="使用中"
			elseif  rsz("flog")=0 then
				Myarray(k)="空闲"
			end if
	rsz.movenext
	k=k+1
	loop
	call closeRs(rsz)		
	%>
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr>
    <td colspan="5" align="center">会议室情况一览表</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>名称</td>
    <td>状态</td>
    <td>位置</td>
    <td>设施</td>
    <td>容纳人数</td>
  </tr>
  <tr>
    <td>1号会议室</td>
    <td><%= Myarray(0) %>&nbsp;</td>
    <td>1号厂房一楼</td>
    <td rowspan="4"><p>活动电子白板三块</p>
        <p>活动投影仪二台</p></td>
    <td>6人</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>2号会议室</td>
    <td><%= Myarray(1) %></td>
    <td>1号厂房模具夹层二楼</td>
    <td>8人</td>
  </tr>
  <tr>
    <td>3号会议室</td>
    <td><%= Myarray(2) %></td>
    <td>1号厂房模具夹层二楼</td>
    <td>12人</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>4号会议室</td>
    <td><%= Myarray(3) %></td>
    <td>1号厂房模具夹层二楼</td>
    <td>12人</td>
  </tr>
  <tr>
    <td>5号会议室</td>
    <td><%= Myarray(4) %></td>
    <td>1号厂房二楼</td>
    <td><p>活动电子白板一块</p>
        <p>固定投影仪一台</p></td>
    <td>14人</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>6号会议室</td>
    <td><%= Myarray(5) %></td>
    <td>2号厂房成形办公室三楼</td>
    <td rowspan="2"><p>活动电子白板一块</p>
        <p>活动投影仪一台</p></td>
    <td>8人</td>
  </tr>
  <tr>
    <td>7号会议室</td>
    <td><%= Myarray(6) %></td>
    <td>2号厂房成形办公室三楼</td>
    <td>6人</td>
  </tr>
  <tr>
    <td>8号多功能厅</td>
    <td><%= Myarray(7) %></td>
    <td>多功能厅</td>
    <td>&nbsp;</td>
	<td>100人</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="5" align="center"><img src="../img/map.gif" width="400" height="200" /></td>
  </tr>
</table>
