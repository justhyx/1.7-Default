<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
Dim DBPath,ConnStr,Conn,Rs,Sql
	ConnStr="PROVIDER=SQLOLEDB;DATA SOURCE=(local);database=hudsonwwwroot;uid=sa;pwd="
	Set Conn=Server.CreateObject("ADODB.Connection")
	Conn.Open ConnStr

sub CreateRs(rs,sql,n1,n2)
	set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn,n1,n2
end sub

sub closeRs(rs)
	rs.close
	set rs=nothing
end sub

sub closeConn()
	conn.close
	set conn=nothing
end sub
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="refresh" content="5" >
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-color: #E5E2DA;
}
.STYLE1 {
	color: #000000;
	font-weight: bold;
}
-->
</style></head>

<body>
				<% 
				dim s,j,k,l,m,n,p
				sa=0
				ja=0
				ka=0
				la=0
				ma=0
				na=0
				pa=0
				su=0
				ju=0
				ku=0
				lu=0
				mu=0
				nu=0
				pu=0
				si=0
				ji=0
				ki=0
				li=0
				mi=0
				ni=0
				pi=0
				call CreateRs(rs,"select * from UseCar where Urgencyfloy='"&int(2)&"'",1,1)
					do while not rs.eof
						if rs("CarDepartment")="部品技术课" or rs("CarDepartment")="技术课" then 
							ja=ja+1
						elseif rs("CarDepartment")="人事总务课" or rs("CarDepartment")="生产管理课" or rs("CarDepartment")="财务课" then
							sa=sa+1
						elseif rs("CarDepartment")="营业课" or rs("CarDepartment")="营业采购课" then
							ka=ka+1
						elseif rs("CarDepartment")="仓库管理课" then
							la=la+1
						elseif rs("CarDepartment")="成形课" then
							ma=ma+1
						elseif rs("CarDepartment")="注塑检查课" then
							na=la+1
						elseif rs("CarDepartment")="生产课" then
							pa=pa+1
						end if
						rs.movenext
						loop
						call closeRs(rs)						
				call CreateRs(rs,"select * from UseCar where arrange='"&int(1)&"'",1,1)	
					do while not rs.eof	
						if rs("CarDepartment")="部品技术课" or rs("CarDepartment")="技术课" then 
							ju=ju+1
						elseif rs("CarDepartment")="人事总务课" or rs("CarDepartment")="生产管理课" or rs("CarDepartment")="财务课" then
							su=su+1
						elseif rs("CarDepartment")="营业课" or rs("CarDepartment")="营业采购课" then
							ku=ku+1
						elseif rs("CarDepartment")="仓库管理课" then
							lu=lu+1
						elseif rs("CarDepartment")="成形课" then
							mu=mu+1
						elseif rs("CarDepartment")="注塑检查课" then
							nu=lu+1
						elseif rs("CarDepartment")="生产课" then
							pu=pu+1
						end if
					rs.movenext
					loop
				call closeRs(rs)
			call closeConn()				
				 %>
				<table width="396" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
                  <tr>
                    <td align="center" bgcolor="#E5E2DA"><span class="STYLE1">待处理事项</span> </td>
                  </tr>
                  <tr>
                    <td bgcolor="#E5E2DA">&nbsp;</td>
                  </tr>
</table>
				<table width="396" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
                  <tr align="center">
                    <td bgcolor="#E5E2DA">事项\部门</td>
                    <td bgcolor="#E5E2DA">技术</td>
                    <td bgcolor="#E5E2DA">总务</td>
                    <td bgcolor="#E5E2DA">营业</td>
                    <td bgcolor="#E5E2DA">仓库</td>
                    <td bgcolor="#E5E2DA">成形</td>
                    <td bgcolor="#E5E2DA">品检</td>
                    <td bgcolor="#E5E2DA">生产</td>
                  </tr>
                  <tr>
                    <td align="right" bgcolor="#E5E2DA">急用车</td>
                    <td align="center" bgcolor="#E5E2DA"><%= sa %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= ja %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= ka %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= la %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= ma %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= na %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= pa %>&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="right" bgcolor="#E5E2DA">用　车</td>
                    <td align="center" bgcolor="#E5E2DA"><%= ju %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= su %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= ku %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= lu %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= mu %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= nu %>&nbsp;</td>
                    <td align="center" bgcolor="#E5E2DA"><%= pu %>&nbsp;</td>
                  </tr>
</table>
</body>
</html>
