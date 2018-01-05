<!--#include file="../connet/conn.asp" -->
<% 
action=Trim(Request.QueryString("action"))
select case action
	case "add"
		call add()
	case "modi"
		call modi()
	case "dell"
		call dell()
	case "is_ok"
		call is_ok()
	case "manager_ok"
		call manager_ok()
	case "group_ok"
		call group_ok()
end select
	sub add()
	if Trim(Request.Form("department"))="" or Trim(Request.Form("start_1"))="" or Trim(Request.Form("start_2"))="" or Trim(Request.Form("over_1"))="" or Trim(Request.Form("over_2"))="" or Trim(Request.Form("note"))="" then
	response.Write"<script>alert('请将信息填写完整！'); history.back();</script>"
	response.end()
	end if
	call CreateRs(rs,"select * from overtime",1,3)
	rs.addnew
	rs("department")=Trim(Request.Form("department"))
	rs("user_l")=session("user_l")
	rs("over_l")=session("over_l")
	rs("start_time_1")=Trim(Request.Form("start_1"))
	rs("start_time_2")=Trim(Request.Form("start_1"))&" "&Trim(Request.Form("start_2"))
	rs("over_time_1")=Trim(Request.Form("over_1"))
	rs("over_time_2")=Trim(Request.Form("over_1"))&" "&Trim(Request.Form("over_2"))
	rs("note")=Trim(Request.Form("note"))&"预算外"
	rs("username")=session("UserName")
	rs("mm")=month(Trim(Request.Form("start_1")))
	rs("manager_ok")=1
	rs("other_flog")=1
	rs("apply_time")=now()	
	d1=Trim(Request.Form("start_1")&" "&Request.Form("start_2"))
	d2=Trim(Request.Form("over_1")&" "&Request.Form("over_2"))
	overtime=datediff("n",d1,d2)/60
	xdtime=datediff("n",d1,now())/60
	if overtime<-23 or overtime>23 then
		response.Write"<script> alert('日期选择不正确，请仔细查看或联系管理员。管理员短号6538'); history.back();</script>"
		response.End()
	end if
	if xdtime>24 then
		response.Write"<script> alert('超过24小时不能系统填写加班申请，请手写加班申请'); history.back();</script>"
		response.End()
	end if		
	ww=Trim(weekday(Request.Form("start_1")))
	if ww=1 or ww=7 then
		rs("isweek")=Formatnumber(overtime)
		rs("noweek")=0		
	 else
	 	rs("noweek")=Formatnumber(overtime)
		rs("isweek")=0
	end if	
	rs.update
	call closeRs(rs)
	response.Redirect("other_add.asp")
	end sub
	sub modi()
	if Trim(Request.Form("department"))="" or Trim(Request.Form("start_1"))="" or Trim(Request.Form("start_2"))="" or Trim(Request.Form("over_1"))="" or Trim(Request.Form("over_2"))="" or Trim(Request.Form("note"))="" then
	response.Write"<script>alert('请将信息填写完整！'); history.back();</script>"
	response.end()
	end if
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("department")=Trim(Request.Form("department"))
	rs("start_time_1")=Trim(Request.Form("start_1"))
	rs("start_time_2")=Trim(Request.Form("start_2"))
	rs("over_time_1")=Trim(Request.Form("over_1"))
	rs("over_time_2")=Trim(Request.Form("over_2"))
	rs("mm")=month(Trim(Request.Form("start_1")))
	rs("note")=Trim(Request.Form("note"))
	rs("username")=session("UserName")
	d1=Trim(Request.Form("start_1")&" "&Request.Form("start_2"))
	d2=Trim(Request.Form("over_1")&" "&Request.Form("over_2"))
	overtime=datediff("n",d1,d2)	
	ww=Trim(weekday(Request.Form("start_1")))
	if ww=1 or ww=7 then
		rs("isweek")=Formatnumber(overtime/60)
		rs("noweek")=0		
	 else
	 	rs("noweek")=Formatnumber(overtime/60)
		rs("isweek")=0
	end if	
	rs.update
	call closeRs(rs)
	response.Redirect("other_add.asp")
	end sub
	sub dell()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs.delete
	rs.update
	call closeRs(rs)
	response.Redirect("other_add.asp")
	end sub
	sub is_ok()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("usermanager")=session("UserName")
	rs("is_ok")=Trim(Request.QueryString("is_ok"))
	rs.update
	call closeRs(rs)
	response.Redirect("other_add.asp")
	end sub
	sub manager_ok()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("manager_ok")=0
	rs("usermanager")=session("username")
	rs.update
	call closeRs(rs)
	response.Redirect("list_ot.asp")
	end sub
sub group_ok()
	for i=1 to int(Trim(Request.Form("rs_i")))
		if request.Form("cbox"&i&"")<>"" then
			call CreateRs(rs,"select * from overtime where id="&Trim(Request.Form("cbox"&i&"")),1,3)
				rs("manager_ok")=0
				rs("usermanager")=session("username")	
				rs.update
		end if
	next
		call closeRs(rs)
		response.redirect("list_ot.asp")
		
end sub
call closeConn()
 %>







