<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<style>
li{
	list-style-type:none;
	}
ul{
	font-weight:bold;
	}
	
</style>
<body>
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1">
  <tr bgcolor="#999999">
    <td height="40" colspan="2" align="center" style="font-size:24px; font-weight:bold; color:#FFFFFF;">和诚锴诚公司规定标准</td>
  </tr>
  <tr bgcolor="#999999">
    <td width="259">&nbsp;</td>
    <td width="682">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#CCCCCC" ><ul>
        手册
        <% 
			call createRs(rs,"select * from company where columnid='"&int(1)&"' and folg='"&int(1)&"'",1,1)
			do while not rs.eof
		 %>
        <li>&nbsp;&nbsp;<a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
        <% 
		rs.movenext
		loop
		call closeRs(rs)
		 %>
      </ul>
      <ul>
        标准
        <% 
			call createRs(rs,"select * from company where columnid='"&int(2)&"' and folg='"&int(1)&"'",1,1)
			do while not rs.eof
		 %>
        <li>&nbsp;&nbsp;<a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
        <% 
		rs.movenext
		loop
		call closeRs(rs)
		 %>
      </ul>
      <ul>
        规定
        <li>&nbsp;&nbsp;┗ 部品技术课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(4)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 仓库管理课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(6)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 人事总务课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(9)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 生产管理课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(7)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 营业采购课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(5)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 注塑检查课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(8)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 财务课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(12)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 成形课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(10)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 技术课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(13)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
        <li>&nbsp;&nbsp;┗ 生产课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='"&int(11)&"' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
		 <li>&nbsp;&nbsp;┗ IT 技术课
          <ul>
            <% 
					call createRs(rs,"select * from company where columnid='e' and folg='"&int(1)&"'",1,1)
					do while not rs.eof
				 %>
            <li><a href="companylist.asp?content=<%= rs("columnurl") %>">├<%= rs("columncontent") %></a></li>
            <% 
				rs.movenext
				loop
				call closeRs(rs)
				 %>
          </ul>
        </li>
      </ul></td>
    <td valign="top"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="960" height="600">
        <param name="movie" value="../companycontent/<%= Trim(Request.QueryString("content")) %>" />
        <param name="quality" value="high" />
        <embed src="../companycontent/<%= Trim(Request.QueryString("content")) %>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="550" height="388"></embed>
      </object></td>
  </tr>
</table>
</body>
</html>
<% call closeConn() %>
