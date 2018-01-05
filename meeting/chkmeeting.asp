<!--#include file="../connet/conn.asp" -->
<% 
 dim flog
 if Trim(Request.Form("action"))="add" then
 call CreateRs(rs,"select * from meeting",1,3)
 	rs.addnew
	rs("M_Department")=Trim(Request.Form("M_Department"))
	rs("M_usename")=Trim(Request.Form("M_usename"))
	rs("M_usetime")=Trim(Request.Form("M_usetime"))
	rs("M_overtime")=Trim(Request.Form("M_overtime"))
	rs("M_guest")=Trim(Request.Form("M_guest"))
	rs("M_number")=Trim(Request.Form("M_number"))
	rs("M_device")=Trim(Request.Form("M_device"))
	rs("M_usetime_h")=Trim(Request.Form("M_usetime_h"))
	rs("M_overtime_h")=Trim(Request.Form("M_overtime_h"))
	rs("M_device")=Trim(Request.Form("M_device"))
	rs("M_device1")=Trim(Request.Form("M_device1"))
	rs("intention")=Trim(Request.Form("intention"))
	rs("addtime")=now()
	rs("M_flog")="等待安排"		
	rs.update
 call closeRs(rs)
 response.Redirect("foxmail.asp")
 elseif Trim(Request.Form("action"))="update" then
 call CreateRs(rs,"select * from meeting where id="&Trim(Request.Form("id")),1,3)
 rs("M_flog")=Trim(Request.Form("M_flog"))
 rs("M_meetingroom")=Trim(Request.Form("M_meetingroom"))
 rs("M_manger")=session("UserName")
 rs("uptime")=now() 
 rs.update 
 call CreateRs(rsc,"select * from meetroom where id="&int(Trim(Request.Form("M_meetingroom"))),1,3)
 select case Trim(Request.Form("M_flog"))	
 	case "预订完成" 
		rsc("flog")=2
	case "使用中"
		rsc("flog")=1
	case "使用结束"
		rsc("flog")=0
	end select
	rsc.update
	call closeRs(rsc)
 end if  
 response.Redirect("meetinglist.asp")
 call closeRs(rs)
 call closeConn()
 %>