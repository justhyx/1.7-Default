<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
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
      <td>������鿴�·ݣ�
        <label>
        <input name="mmlist" type="text" id="mmlist" />
        <input type="submit" name="Submit" value="ȷ ��" />
      ��Ĭ��Ϊ��ǰ�£� | ��ǰ�鿴�·ݣ� <%= session("mmlist") %></label></td>
    </tr>
  </table>
</form>
<table width="100%" border="0" cellspacing="1" cellpadding="4">
    <tr>
    <td bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">Ʒ�ʼ���</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='Ʒ�ʼ���' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">��Ʒ������ </td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='��Ʒ������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">�����</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">Ӫҵ�ɹ���</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='Ӫҵ�ɹ���' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���������</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='���������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������</td>
  </tr>
  <% 
  'call CreateRs(rs,"select * from HUDSON_User where department='������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���μ�����</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='���μ�����' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���ο�</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='���ο�' and user_l<>'3' and user_l<>'4' and user_l<>'5' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���ο�--G1</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���ο�--G2</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���ο�--G3</td>
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
    <td  bgcolor="00CCFF" style="font-weight:bold; color:#FFFFFF;">������</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������-ϳ��</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������-ĥ��</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������-CNC</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������-�ŵ�</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������-�߸�</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">������-ʡģ&amp;����</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">��������</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">��������� ����</td>
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">�����</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">���������</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='���������' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">IT������</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='IT��' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">�ξ߿�</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='�ξ߿�' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">�豸</td>
  </tr>
  <% 
  call CreateRs(rs,"select * from HUDSON_User where department='�豸��' order by username",1,1)
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
    <td  bgcolor="#00CCFF" style="font-weight:bold; color:#FFFFFF;">ISO�ƽ�</td>
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






