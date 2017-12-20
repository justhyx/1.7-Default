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
call createRs(rs,"select * from equip where ipdress='"&jt&"'",1,1)
jtbh=rs("equip_id")
jt_name=rs("equip_name")
call closeRs(rs)
Dim TxtKey  '选择计划
TxtKey=Trim(Request.QueryString("TxtKey"))
If TxtKey="" Then
	call CreateRs(rs,"select * from prd_dictate where update_state='N' and equip_id='"&jtbh&"' and dictate_date='"&date()&"'",1,1)
		if rs.eof or rs.bof then
			response.Write("没有计划")
			response.End()
		else
			call CreateRs(rsa,"select * from goods where goods_id='"&rs("goods_id")&"'",1,1)
			bf=rsa("goods_name")
			cxzq=rsa("cxzq")
			qs=rsa("qs")
			response.cookies("prd_dictate_id")=rs("dictate_id")
			response.cookies("jhs")=rs("prd_qty")
			response.Cookies("jhs").Expires=dateadd("h",24,now())
			response.Cookies("prd_dictate_id").Expires=dateadd("h",24,now())
			call closeRs(rsa)			
		end if
Else
	call CreateRs(rs,"select * from prd_dictate where dictate_id='"&TxtKey&"'",1,1)
		if rs.eof or rs.bof then
			response.Write("参数错误")
			response.End()
		else
			call CreateRs(rsa,"select * from goods where goods_id='"&rs("goods_id")&"'",1,1)
			bf=rsa("goods_name")
			cxzq=rsa("cxzq")
			qs=rsa("qs")
			response.cookies("jhs")=rs("prd_qty")
			response.Cookies("jhs").Expires=dateadd("h",24,now())
			call closeRs(rsa)
		end if
End If	
 %>
<body>
<form id="form1" name="form1" method="post" action="chk.asp?event=start">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000">
    <tr>
      <td colspan="5" align="center" bgcolor="#FFFFFF" style="font-weight:bold;"><%= date() %>&nbsp;&nbsp;&nbsp;&nbsp;计 划</td>
    </tr>
    <tr>
      <td width="11%" bgcolor="#FFFFFF">计划日期</td>
      <td colspan="4" bgcolor="#FFFFFF"><%= date() %>&nbsp;</td>
    </tr>
    <tr>
      <td rowspan="4" bgcolor="#FFFFFF">计划<%= rs("dictate_BCD_val") %></td>
      <td width="7%" bgcolor="#FFFFFF">担当</td>
      <td width="15%" bgcolor="#FFFFFF"><input name="dd" type="text" id="dd" size="16" /></td>
      <td width="5%" bgcolor="#FFFFFF">班别</td>
      <td width="62%" bgcolor="#FFFFFF"><input name="bb" type="text" id="bb" size="16" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">装箱数</td>
      <td bgcolor="#FFFFFF"><input name="zxs" type="text" id="zxs" size="16" /></td>
      <td bgcolor="#FFFFFF">作业员</td>
      <td bgcolor="#FFFFFF"><input name="zyy" type="text" id="zyy" size="16" /></td>
    </tr>	
    <tr>
      <td bgcolor="#FFFFFF">机台</td>
      <td bgcolor="#FFFFFF"><input name="jt" type="text" id="jt" value="<%= jt_name %>" size="16" /></td>
      <td bgcolor="#FFFFFF">计划数</td>
      <td bgcolor="#FFFFFF"><input name="js" type="text" id="js" value="<%= rs("prd_qty") %>" size="16" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">部番</td>
	  <td bgcolor="#FFFFFF"><input name="bf" type="text" id="bf" value="<%= bf %>" size="16" /></td>
	  <td bgcolor="#FFFFFF">周期</td>
      <td bgcolor="#FFFFFF"><input name="zq" type="text" id="zq" value="<%= cxzq %>" size="16" />
      <input name="qs" type="hidden" id="qs" value="<%= qs %>" />
      取数：<%= qs %></td>
	  <% call closeRs(rs) %>
    </tr>
    <% 
		call CreateRs(rs,"select * from prd_dictate where (dictate_date between '"&date()-1&"' and '"&date()+1&"') and update_state='N' and equip_id='"&jtbh&"' order by dictate_BCD_val",1,1)
		if rs.eof then
		response.write"<script> alert('机台信息错误或者没有计划请联系担当确认');</script>"
		response.End()
		end if
		j=0
		do while not rs.eof
			call CreateRs(rsa,"select * from goods where goods_id='"&rs("goods_id")&"'",1,1)
			bf=rsa("goods_name")
			call closeRs(rsa)
				 %>
	<tr>
      <td bgcolor="#FFFFFF"><a href="zs_jh.asp?TxtKey=<%= rs("dictate_id") %>">计划<%= rs("dictate_BCD_val") %></a></td>
      <td bgcolor="#FFFFFF">部番</td>
      <td bgcolor="#FFFFFF"><%= bf %></td>
      <td bgcolor="#FFFFFF">数量</td>
      <td bgcolor="#FFFFFF"><%= rs("prd_qty") %></td>
	</tr>
	<% 
	rs.movenext
	j=j+1
	loop
	 %>
    <% 
		call CreateRs(rs,"select * from prd_dictate where (dictate_date between '"&date()-1&"' and '"&date()+1&"') and update_state='Y' and equip_id='"&jtbh&"'",1,1)
		k=0
		do while not rs.eof
			call CreateRs(rsa,"select * from goods where goods_id='"&rs("goods_id")&"'",1,1)
			bf=rsa("goods_name")
			call closeRs(rsa)			
	 %>
	<tr>
      <td bgcolor="#FFFFFF">计划<%= rs("dictate_BCD_val") %>(完成)</td>
      <td bgcolor="#FFFFFF">部番</td>
      <td bgcolor="#FFFFFF"><%= bf %></td>
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
    <input style="height:40px; width:71px;" type="submit" value="    开 始    " />
</center>
</form>
</body>
</html>
<% 
call closeConn()
 %>