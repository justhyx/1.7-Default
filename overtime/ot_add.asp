<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
	if mid(session("Power"),3,1)="0" then 
		response.redirect("group.asp")
	end if
	gu=Trim(Request.QueryString("op"))
	if gu="" then
		 gudge="add"
		 s_t_1=date()
		 o_t_1=date()
	elseif gu="modi" then
		call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
		gudge="modi"
		 s_t_1=rs("start_time_1")
		 s_t_2=rs("start_time_2")
		 o_t_1=rs("over_time_1")
		 o_t_2=rs("over_time_2")
		 bm=rs("department")
		 why=rs("note")
		 id=rs("id")
		call closeRs(rs)		
	end if	
	%>
</head>

<body>
<form id="form1" name="form1" method="post" action="chk.asp?action=<%= gudge %>&id=<%= id %>">
  <table width="100%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
    <tr>
      <td colspan="5" align="center" style="font-weight:bold;">加班申请单填写</td>
    </tr>
    <tr>
      <td>部门</td>
      <td>开始时间</td>
      <td>结束时间</td>
      <td>辅助</td>
      <td>事由</td>
    </tr>
    <tr>
      <td><label>
        <input name="department" type="text" id="department" value="<%= session("Department")%>" size="20" />
        <input name="username" type="hidden" id="username" value="<%= session("UserName") %>" />
      </label></td>
      <td><label>
        <input name="start_1" type="text" id="start_1" value="<%= s_t_1%>" size="9" onblur="if(/^[0-9]{4}-((([13578]|(10|12))-([1-9]|[1-2][0-9]|3[0-1]))|(2-([1-9]|[1-2][0-9]))|(([469]|11)-([1-9]|[1-2][0-9]|30)))$/.test(start_1.value)==false) {alert('时间格式不正确\n例：2012-7-20'); return false;}" />
        <select name="start_2" id="start_2">
          <option value="0:00">00:00</option>
          <option value="0:30">00:30</option>
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
        <input name="over_1" type="text" id="over_1" value="<%= o_t_1 %>" size="9"  onblur="if(/^[0-9]{4}-((([13578]|(10|12))-([1-9]|[1-2][0-9]|3[0-1]))|(2-([1-9]|[1-2][0-9]))|(([469]|11)-([1-9]|[1-2][0-9]|30)))$/.test(over_1.value)==false) {alert('时间格式不正确\n例：2012-7-20'); return false;}" />
        <select name="over_2" id="over_2">
          <option value="0:00">00:00</option>
          <option value="0:30">00:30</option>
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
       (<font color="#FF0000"><strong>自动减去12:00--13:30的1.5小时中午休息时间</strong></font>) <br />
      <input name="midday" type="checkbox" id="midday" value="b" />
      中午连班(<strong><font color="#FF0000">自动加上中午连班1个小时 勾选系统自动事由填写中午连班 </font></strong>)<br />
      <input name="after" type="checkbox" id="after" value="c" />
      减去下午17:30--18:00半小时
      </label></td>
      <td><input name="note" type="text" id="note" value="<%= why%>" size="50" /></td>
    </tr>
    <tr>
      <td colspan="5" align="center"><label>
        <input type="submit" name="Submit" value="提 交" />
      </label></td>
    </tr>
  </table>
</form>
<table width="100%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr style="font-weight:bold;">
    <td>部门</td>
    <td>开始时间</td>
    <td>结束时间</td>
    <td>事由</td>
    <td>非周末</td>
    <td>周末</td>	
    <td>操作</td>
  </tr>
  <% 
  	call CreateRs(rs,"select * from overtime where username='"&session("UserName")&"' and other_flog='"&int(0)&"' and manager_ok='"&int(1)&"' order by start_time_1",1,1)
	if rs.eof then 
	response.End()
	end if
	is_week=0
	no_week=0
	do while not rs.eof
	d1=rs("start_time_2")
	d2=rs("over_time_2")
	overtime=datediff("n",d1,d2) 
	%>
  <tr>
    <td><%= rs("department") %>&nbsp;</td>
    <td><%= rs("start_time_1")%>&nbsp;<%= rs("start_time_2")%></td>
    <td><%= rs("over_time_1") %>&nbsp;<%= rs("over_time_2") %></td>
    <td><%= rs("note") %></td>
    <td><%=rs("noweek")  %></td>
    <td><%=rs("isweek")  %></td>
    <td><a href="ot_add.asp?op=modi&id=<%= rs("id") %>">修改</a> | <a href="chk.asp?action=dell&amp;id=<%= rs("id") %>">取消</a></td>
  </tr>
  <% 
    is_week=rs("isweek")+is_week
  	no_week=rs("noweek")+no_week
  rs.movenext
  loop
   %>
  <tr>
    <td colspan="4" align="right">小计：</td>
    <td><%= no_week %></td>
    <td><%= is_week %></td>
    <td>总计：<%= is_week + no_week %></td>
  </tr>
  <tr>
    <td colspan="7"><%=session("UserName")  %>&nbsp;</td>
  </tr>
</table>
<table width="100%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr style="font-weight:bold;">
    <td>部门</td>
    <td>开始时间</td>
    <td>结束时间</td>
    <td>事由</td>
    <td>非周末</td>
    <td>周末</td>
    <td>备注</td>
    <td>操作</td>
  </tr>
  <% 
  	call CreateRs(rs,"select * from overtime where username='"&session("UserName")&"' and other_flog='"&int(0)&"' and manager_ok='"&int(2)&"' order by start_time_1",1,1)
	if rs.eof then 
	response.End()
	end if
	is_week=0
	no_week=0
	do while not rs.eof
	d1=rs("start_time_2")
	d2=rs("over_time_2")
	overtime=datediff("n",d1,d2) 
	%>
  <tr>
    <td><%= rs("department") %>&nbsp;</td>
    <td><%= rs("start_time_1")%>&nbsp;<%= rs("start_time_2")%></td>
    <td><%= rs("over_time_1") %>&nbsp;<%= rs("over_time_2") %></td>
    <td><%= rs("note") %></td>
    <td><%=rs("noweek")  %></td>
    <td><%=rs("isweek")  %></td>
    <td><font color="#FF0000">被取消加班</font></td>
    <td><a href="chk.asp?action=dell&amp;id=<%= rs("id") %>">确 定</a></td>
  </tr>
  <% 
    is_week=rs("isweek")+is_week
  	no_week=rs("noweek")+no_week
  rs.movenext
  loop
   %>
  <tr>
    <td colspan="4" align="right">小计：</td>
    <td><%= no_week %></td>
    <td><%= is_week %></td>
    <td>总计：<%= is_week + no_week %></td>
    <td>&nbsp;</td>
  </tr>  
</table>
<% 
call closeRs(rs)
call closeConn()
 %>
</body>
</html>







