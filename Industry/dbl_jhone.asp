<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>和诚锴诚工业系统--生产计划</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<%
jt=Request.ServerVariables("REMOTE_ADDR")
call createRs(rs,"select * from 机台绑定 where IPD_ip='"&jt&"'",1,1)
jtbh=rs("ji")
call closeRs(rs)
Dim TxtKey
TxtKey=Trim(Request.QueryString("TxtKey"))
If TxtKey="" Then
	call CreateRs(rs,"select * from 计划白板 where state='N' and equip_name='"&jtbh&"'",1,1)
		if rs.eof or rs.bof then
			response.Write("没有计划")
			response.End()
		end if
Else
	call CreateRs(rs,"select * from 计划白板 where state='N' and equip_name='"&jtbh&"' and id='"&TxtKey&"'",1,1)
		if rs.eof or rs.bof then
			response.Write("参数错误")
			response.End()
		end if
End If	
 %>
<body>
<form id="form1" name="form1" method="post" action="chk_one.asp?event=start">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000">
    <tr>
      <td colspan="5" align="center" bgcolor="#FFFFFF" style="font-weight:bold;"><%= date() %>&nbsp;&nbsp;&nbsp;&nbsp;计 划</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">计划日期</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= date() %></td>
    </tr>
    <tr>
      <td rowspan="4" bgcolor="#FFFFFF">计划<%= rs("dictate_no") %></td>
      <td bgcolor="#FFFFFF">部番</td>
      <td bgcolor="#FFFFFF"><input name="bf" type="text" id="bf" value="<%= rs("goods_name") %>" size="16" /></td>
      <td bgcolor="#FFFFFF">班别</td>
      <td bgcolor="#FFFFFF"><input name="bb" type="text" id="bb" size="16" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">计划数</td>
      <td bgcolor="#FFFFFF"><input name="js" type="text" id="js" value="<%= rs("prd_qty") %>" size="16" /></td>
      <td bgcolor="#FFFFFF">作业员</td>
      <td bgcolor="#FFFFFF"><input name="zyy" type="text" id="zyy" size="16" /></td>
    </tr>	
    <tr>
      <td bgcolor="#FFFFFF">机台</td>
      <td bgcolor="#FFFFFF"><input name="jt" type="text" id="jt" value="<%= jtbh %>" size="16" /></td>
      <td bgcolor="#FFFFFF">装箱数</td>
      <td bgcolor="#FFFFFF"><input name="zxs" type="text" id="zxs" size="16" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">周期</td>
	  <td bgcolor="#FFFFFF"><input name="zq" type="text" id="zq" value="<%= rs("cxzq") %>" size="16" /></td>
	  <td bgcolor="#FFFFFF">担当</td>
      <td bgcolor="#FFFFFF"><input name="dd" type="text" id="dd" size="16" />
      <input name="qs" type="hidden" id="qs" value="<%= rs("qs") %>" /></td>
	  <% call closeRs(rs) %>
    </tr>
    <% 
		call CreateRs(rs,"select * from 计划白板 where state='N' and equip_name='"&jtbh&"'",1,1)
		if rs.eof then
		response.write("机台信息错误或者没有计划请联系担当确认")
		response.End()
		end if
		j=0
		do while not rs.eof		
	 %>
	<tr>
      <td width="120" bgcolor="#FFFFFF"><a href="dbl_jhone.asp?TxtKey=<%= rs("id") %>">计划<%= rs("dictate_no") %></a></td>
      <td width="70" bgcolor="#FFFFFF">部番</td>
      <td width="120" bgcolor="#FFFFFF"><%= rs("goods_name") %></td>
      <td width="70" bgcolor="#FFFFFF">数量</td>
      <td width="120" bgcolor="#FFFFFF"><%= rs("prd_qty") %></td>
	</tr>
	<% 
	rs.movenext
	j=j+1
	loop
	 %>
    <% 
		call CreateRs(rs,"select * from 计划白板 where state='Y' and equip_name='"&jtbh&"'",1,1)
		k=0
		do while not rs.eof		
	 %>
	<tr>
      <td bgcolor="#FFFFFF">计划<%= rs("dictate_no") %>(完成)</td>
      <td bgcolor="#FFFFFF">部番</td>
      <td bgcolor="#FFFFFF"><%= rs("goods_name") %></td>
      <td bgcolor="#FFFFFF">数量</td>
      <td bgcolor="#FFFFFF"><%= rs("prd_qty") %></td>
	</tr>
	<% 
	rs.movenext
	k=k+1
	loop
	 %>
  </table>
 <center><br/>
    <input style="width:71px;" name="" type="submit" value="    开 始    " />
</center>
</form>
</body>
</html>
<% 
call closeConn()
 %>