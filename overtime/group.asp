<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<% 
	if mid(session("Power"),3,1)<>"0" then		
			response.Write"<script> history.back();</script>"
			response.End()			
	end if
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++){
    var e = form.elements[i];
    if (e.name != 'chkall')
	e.checked = form.chkall.checked; 
   }
  }
</script>
</head>

<body>
<form id="form" name="form" method="post" action="groupchk.asp?action=add">
  <table width="100%" border="1" cellpadding="4" cellspacing="0" bordercolor="#000000">
    <tr>
      <td>担当：</td>
      <td><%= session("UserName") %>&nbsp;</td>
      <td align="right">&nbsp;</td>
      <td>辅助</td>
      <td>事由</td>
    </tr>
    <tr>
      <td><label>加班时间：</label></td>
      <td><label>
        <input name="start_1" type="text" id="start_1" value="<%= date()%>" size="9" onblur="if(/^[0-9]{4}-((([13578]|(10|12))-([1-9]|[1-2][0-9]|3[0-1]))|(2-([1-9]|[1-2][0-9]))|(([469]|11)-([1-9]|[1-2][0-9]|30)))$/.test(start_1.value)==false) {alert('时间格式不正确\n例：2012-7-20'); return false;}" />
        <select name="start_2" id="start_2">
          <option value="00:00">00:00</option>
          <option value="00:30">00:30</option>
          <option value="1:00">1:00</option>
          <option value="1:30">1:30</option>
          <option value="2:00">2:00</option>
          <option value="2:30">2:30</option>
          <option value="3:00">3:00</option>
          <option value="3:30">3:30</option>
          <option value="4:00">4:00</option>
          <option value="4:30">4:30</option>
          <option value="5:00">5:00</option>
          <option value="5:30">5:30</option>
          <option value="6:00">6:00</option>
          <option value="6:30">6:30</option>
          <option value="7:00">7:00</option>
          <option value="7:30">7:30</option>
          <option value="8:00">8:00</option>
          <option value="8:30">8:30</option>
          <option value="9:00">9:00</option>
          <option value="9:30">9:30</option>
          <option value="10:00">10:00</option>
          <option value="10:30">10:30</option>
          <option value="11:00">11:00</option>
          <option value="11:30">11:30</option>
          <option value="12:00">12:00</option>
		  <option value="12:15">12:15</option>
          <option value="12:30">12:30</option>
          <option value="13:00">13:00</option>
		  <option value="13:15">13:15</option>
          <option value="13:30">13:30</option>
          <option value="14:00">14:00</option>
          <option value="14:30">14:30</option>
          <option value="15:00">15:00</option>
          <option value="15:30">15:30</option>
          <option value="16:00">16:00</option>
          <option value="16:30">16:30</option>
          <option value="17:00">17:00</option>
          <option value="17:30">17:30</option>
          <option value="18:00" selected="selected">18:00</option>
          <option value="18:30">18:30</option>
          <option value="19:00">19:00</option>
          <option value="19:30">19:30</option>
          <option value="20:00">20:00</option>
          <option value="20:30">20:30</option>
          <option value="21:00">21:00</option>
          <option value="21:30">21:30</option>
          <option value="22:00">22:00</option>
          <option value="22:30">22:30</option>
          <option value="23:00">23:00</option>
          <option value="23:30">23:30</option>
        </select>
      </label></td>
      <td><label>
        <input name="over_1" type="text" id="over_1" value="<%= date() %>" size="9"  onblur="if(/^[0-9]{4}-((([13578]|(10|12))-([1-9]|[1-2][0-9]|3[0-1]))|(2-([1-9]|[1-2][0-9]))|(([469]|11)-([1-9]|[1-2][0-9]|30)))$/.test(over_1.value)==false) {alert('时间格式不正确\n例：2012-7-20'); return false;}" />
        <select name="over_2" id="over_2">
          <option value="00:00">00:00</option>
          <option value="00:30">00:30</option>
          <option value="1:00">1:00</option>
          <option value="1:30">1:30</option>
          <option value="2:00">2:00</option>
          <option value="2:30">2:30</option>
          <option value="3:00">3:00</option>
          <option value="3:30">3:30</option>
          <option value="4:00">4:00</option>
          <option value="4:30">4:30</option>
          <option value="5:00">5:00</option>
          <option value="5:30">5:30</option>
          <option value="6:00">6:00</option>
          <option value="6:30">6:30</option>
          <option value="7:00">7:00</option>
          <option value="7:30">7:30</option>
          <option value="8:00">8:00</option>
          <option value="8:30">8:30</option>
          <option value="9:00">9:00</option>
          <option value="9:30">9:30</option>
          <option value="10:00">10:00</option>
          <option value="10:30">10:30</option>
          <option value="11:00">11:00</option>
          <option value="11:30">11:30</option>
		  <option value="11:45">11:45</option>
          <option value="12:00">12:00</option>
          <option value="12:30">12:30</option>
          <option value="13:00">13:00</option>
		  <option value="13:15">13:15</option>
          <option value="13:30">13:30</option>
          <option value="14:00">14:00</option>
          <option value="14:30">14:30</option>
          <option value="15:00">15:00</option>
          <option value="15:30">15:30</option>
          <option value="16:00">16:00</option>
          <option value="16:30">16:30</option>
          <option value="17:00">17:00</option>
          <option value="17:30">17:30</option>
          <option value="18:00">18:00</option>
          <option value="18:30">18:30</option>
          <option value="19:00">19:00</option>
          <option value="19:30">19:30</option>
          <option value="20:00" selected="selected">20:00</option>
          <option value="20:30">20:30</option>
          <option value="21:00">21:00</option>
          <option value="21:30">21:30</option>
          <option value="22:00">22:00</option>
          <option value="22:30">22:30</option>
          <option value="23:00">23:00</option>
          <option value="23:30">23:30</option>
        </select>
      </label></td>
      <td><label>
        <input name="addweek" type="checkbox" id="addweek" value="a" />
        (<strong><font color="#FF0000">自动减去12:00--13:30午休1.5小时</font></strong>)<br />
      <input name="midday" type="checkbox" id="midday" value="b" />
      中午连班(<strong><font color="#FF0000">自动加上中午连班1小时 勾选系统自动在事由注明连班 </font></strong>)<br />
      <input name="after" type="checkbox" id="after" value="c" />
      自动减去17:30:--18:00半小时
      </label></td>
      <td><input name="note" type="text" id="note" size="50" /></td>
    </tr>
    <tr>
      <td align="right"><label>
        全选
            <input type="checkbox" name="chkall" onclick="CheckAll(this.form)" />
      </label></td>
      <td colspan="2">人员</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="5"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <% 
			call CreateRs(rs,"select * from hudson_user where user_l='"&session("user_l")&"' and power>99 order by xu, username",1,1)
			if rs.eof then
			response.end()
			end if
			fi=0
			do while not rs.eof
		 %>        
		  <td><input name="cbox<%= fi %>" type="checkbox" id="cbox" value="<%= rs("username")%>" /></td>
          <td><%= rs("username") %></td>
		  <% 
			  fi=fi+1
			  rs.movenext
			  if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi=78 or fi=91 or fi=104 or fi=117 or fi=130 or fi=143 or fi=156 or fi=169 or fi=182 or fi=195 then 
			  %>
	    <tr>
			 	<td>&nbsp;</td>
          		<td>&nbsp;</td>
		</tr> 
			 <% 
			  end if			  
			  loop
			  call closeRs(rs)
		   %>		  
        </tr>
      </table></td>
    </tr>
    <tr>
      <td colspan="3"><input name="rs_i" type="hidden" id="rs_i" value="<%= fi %>" /></td>
      <td><label>
        <input type="submit" name="Submit" value="添加人员" />
      </label></td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form><table width="100%" border="1" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td>姓名</td>
    <td>加班开始时间</td>
    <td>加班结束时间</td>
    <td>事由</td>
    <td>非周末</td>
	<td>周末</td>
	<td>操作</td>
  </tr>
<% 
	Call CreateRs(rs,"select * from overtime where user_l='"&session("user_l")&"' and manager_ok='"&int(1)&"' order by start_time_1 asc",1,1)
	if rs.eof then
	response.end()
	end if
	jbi=0
	do while not rs.eof
 %>
  <tr>
    <td><%= rs("username") %>&nbsp;</td>
    <td><%= rs("start_time_1")%>&nbsp;<%= rs("start_time_2")%>&nbsp;</td>
    <td><%= rs("over_time_1") %>&nbsp;<%= rs("over_time_2") %>&nbsp;</td>
    <td><%= rs("note") %>&nbsp;</td> 
	<td><%= rs("noweek") %>&nbsp;</td>
	<td><%= rs("isweek") %>&nbsp;</td>
	<td><a href="groupchk.asp?id=<%= rs("id")%>&action=dell">移除人员</a>&nbsp;</td>
  </tr>
  <% 
	  jbi=jbi+1
	  rs.movenext
	  loop
   %>
  <tr>
    <td colspan="7" align="center" style="font-weight:bold;">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="7" style="font-weight:bold;">加班人数：<%= jbi %></td>
  </tr>  
</table>
<p>
</body>
</html>







