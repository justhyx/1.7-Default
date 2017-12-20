<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<% 
dim mmlist
mmlist=request.Form("mmlist")
if mmlist="" then
	mmlist=month(date)
end if
session("mmlist")=mmlist
 %>
</head>
<body>
<form id="form1" name="form1" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="4">
    <tr>
      <td>请输入查看月份：
        <label>
        <input name="mmlist" type="text" id="mmlist" />
        <input type="submit" name="Submit" value="确 认" />
      （默认为当前月） | 当前查看月份： <%= session("mmlist") %></label></td>
    </tr>
  </table>
</form>
<table width="100%" border="0" cellspacing="1" cellpadding="4">
    <tr>
    <td bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">品质检查课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='品质检查课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">部品技术课 </td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='部品技术课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">财务课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='行政部' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">营业采购课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='营业采购课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产管理课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='生产管理课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  
<!--    <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课</td>
  </tr>
  <% 
  'call CreateRs(rs,"select * from HUDSON_User where department='生产课' order by username",1,1)
'   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% 'do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%'= rs("username") %>"</a><%'= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		'fi=fi+1
'		rs.movenext
'		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi=78 or fi=91 or fi=104 or fi=117 or fi=130 or fi=143 or fi=156 then
		 %>
	  <tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
	  </tr>
		<% 
		'end if
'		loop
'		call closeRs(rs)
		 %>
      </tr>
    </table></td>
  </tr>-->
  <tr>
    <td>&nbsp;</td>
  </tr>
  
   <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">成形技术课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='成形技术课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  
     <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">成形课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='成形课' and user_l<>'3' and user_l<>'4' and user_l<>'5' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi=78 or fi=91 or fi=104 or fi=117 or fi=130 or fi=143 or fi=156 or fi=169 or fi=182 or fi=195 or fi=208 or fi=221 then
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
    <td>&nbsp;</td>
  </tr>
  
       <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">成形课--G1</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where  user_l='3' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi=78 or fi=91 or fi=104 or fi=117 or fi=130 or fi=143 or fi=156 or fi=169 or fi=182 or fi=195 or fi=208 or fi=221 then
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
    <td>&nbsp;</td>
  </tr>
  
       <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">成形课--G2</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='4' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi=78 or fi=91 or fi=104 or fi=117 or fi=130 or fi=143 or fi=156 or fi=169 or fi=182 or fi=195 or fi=208 or fi=221 then
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
    <td>&nbsp;</td>
  </tr>
  
       <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">成形课--G3</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='5' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi=78 or fi=91 or fi=104 or fi=117 or fi=130 or fi=143 or fi=156 or fi=169 or fi=182 or fi=195 or fi=208 or fi=221 then
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
    <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td  bgcolor="00CCFF" style="font-weight:bold; color:#FFFFFF;">技术课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='技术课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课-铣床</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='37' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课-磨床</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='25' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课-CNC</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='1' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课-放电</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='17' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课-线割</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='38' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产课-省模&amp;组立</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='32' or user_l='44' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">材料再生</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='14' or user_l='10' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
    <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">生产管理课 生管</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where user_l='13' or user_l='30' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 or fi=52 or fi=65 or fi= 78 or fi=91 or fi= 104 or fi= 117 then
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
    <td>&nbsp;</td>
  </tr>
    <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">总务课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='行政部' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
 
   <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">物流管理课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='物流管理课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">IT技术课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='IT课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">治具课</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='治具课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
    <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">设备</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='设备课' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">ISO推进</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='ISO' order by username",1,1)
   fi=0
   %>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <% do while not rs.eof %>
		<td><a href="hrlist.asp?pasname=<%= rs("username") %>"</a><%= rs("username") %>&nbsp;</td>
        <td>&nbsp;</td>
		<% 
		fi=fi+1
		rs.movenext
		if fi=13 or fi=26 or fi=39 then
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
    <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<% 
call closeConn()
 %>






