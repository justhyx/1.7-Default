<!--#include file="../connet/conn.asp" -->
<% 
action=Trim(Request.QueryString("action"))
if session("user_l")="" or session("UserName")="" then
response.Write"<script> alert('登录已超时，请退出重新登录'); location.href= '../login.asp '</script>"
response.End()
end if
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
	case "manager_NG"
		call manager_NG()
	case "group_ok"
		call group_ok()
	case "tryadd"
	    call tryadd()
end select
	sub add()
	if Trim(Request.Form("department"))="" or Trim(Request.Form("start_1"))="" or Trim(Request.Form("start_2"))="" or Trim(Request.Form("over_1"))="" or Trim(Request.Form("over_2"))="" or Trim(Request.Form("note"))="" then
	response.Write"<script>alert('请将信息填写完整！'); history.back();</script>"
	response.end()
	end if
	if left(Request.Form("start_2"),len(Trim(Request.Form("start_2")))-3) < 6 then
	date_t=DateAdd("d",-1,Trim(Request.Form("start_1")))
	date_t1=DateAdd("d",-1,Trim(Request.Form("over_1")))
	else
	date_t=Trim(Request.Form("start_1"))
	date_t1=Trim(Request.Form("over_1"))
	end if	
	call CreateRs(rs,"select * from overtime",1,3)
	rs.addnew
	rs("department")=Trim(Request.Form("department"))
	rs("user_l")=session("user_l")
	rs("over_l")=session("over_l")
	rs("start_time_1")=date_t
	rs("start_time_2")=date_t&" "&Trim(Request.Form("start_2"))
	rs("over_time_1")=date_t1
	rs("over_time_2")=date_t1&" "&Trim(Request.Form("over_2"))
	remark=Trim(Request.Form("note"))
	rs("username")=Trim(Request.Form("username"))
	rs("mm")=month(Trim(Request.Form("start_1")))
	rs("manager_ok")=1
	rs("other_flog")=0	
	rs("apply_time")=now()
	d1=date_t&" "&Request.Form("start_2")
	d2=date_t1&" "&Request.Form("over_2")
	overtime=datediff("n",d1,d2)/60	
	if left(Request.Form("start_2"),len(Trim(Request.Form("start_2")))-3) < 6 then
	xdtime=datediff("n",d1,dateadd("d",-1,now()))/60
	else
	xdtime=datediff("n",d1,now())/60	
	end if
	if Trim(Request.Form("addweek"))="a" then
	overtime=overtime - 1.5
	end if	
	if Trim(Request.Form("midday"))="b" then
	overtime=overtime + 1
	remark=remark&"_中午连班"
	end if
	if Trim(Request.Form("after"))="c" then
	overtime=overtime - 0.5
	end if	
	ww=Trim(weekday(Request.Form("start_1")))
	if overtime<0 or overtime>24 then
		response.Write"<script> alert('日期选择不正确，请仔细查看或联系管理员。管理员短号6538'); history.back();</script>"
		response.End()
	end if
	if xdtime>24 then
		response.Write"<script> alert('超过24小时不能系统填写加班申请，请手写加班申请'); history.back();</script>"
		response.End()
	end if
	if ww=1 or ww=7 then
			rs("isweek")= overtime
			rs("noweek")=0		
		 else	 	
			rs("noweek")=overtime
			rs("isweek")=0
	end if
	rs("note")=remark
	rs.update
	call closeRs(rs)
	response.Redirect("ot_add.asp")
	end sub
	sub modi()
	if Trim(Request.Form("department"))="" or Trim(Request.Form("start_1"))="" or Trim(Request.Form("start_2"))="" or Trim(Request.Form("over_1"))="" or Trim(Request.Form("over_2"))="" or Trim(Request.Form("note"))="" then
	response.Write"<script>alert('请将信息填写完整！'); history.back();</script>"
	response.end()
	end if
	if left(Request.Form("start_2"),len(Trim(Request.Form("start_2")))-3) < 6 then
	date_t=DateAdd("d",-1,Trim(Request.Form("start_1")))
	date_t1=DateAdd("d",-1,Trim(Request.Form("over_1")))
	else
	date_t=Trim(Request.Form("start_1"))
	date_t1=Trim(Request.Form("over_1"))
	end if	
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("department")=Trim(Request.Form("department"))
	rs("start_time_1")=date_t
	rs("start_time_2")=Trim(Request.Form("start_2"))
	rs("over_time_1")=date_t1
	rs("over_time_2")=Trim(Request.Form("over_2"))
	rs("mm")=month(Trim(Request.Form("start_1")))
	rs("note")=Trim(Request.Form("note"))
	rs("username")=session("UserName")
	d1=Trim(Request.Form("start_1")&" "&Request.Form("start_2"))
	d2=Trim(Request.Form("over_1")&" "&Request.Form("over_2"))
	overtime=datediff("n",d1,d2)/60	
	if left(Request.Form("start_2"),len(Trim(Request.Form("start_2")))-3) < 6 then
	xdtime=datediff("n",d1,dateadd("d",-1,now()))/60
	else
	xdtime=datediff("n",d1,now())/60	
	end if
	if overtime<0 or overtime>24 then
		response.Write"<script> alert('日期选择不正确，请仔细查看或联系管理员。管理员短号6538'); history.back();</script>"
		response.End()
	end if
	if xdtime>23 then
		response.Write"<script> alert('超过24小时不能系统填写加班申请，请手写加班申请'); history.back();</script>"
		response.End()
	end if
	if Trim(Request.Form("addweek"))="a" then
	overtime=overtime - 1.5
	end if
	if Trim(Request.Form("midday"))="b" then
	overtime=overtime + 1
	end if
	if Trim(Request.Form("after"))="c" then
	overtime=overtime - 0.5
	end if	
	ww=Trim(weekday(Request.Form("start_1")))	
	if ww=1 or ww=7 then
			rs("isweek")= overtime
			rs("noweek")=0		
		 else	 	
			rs("noweek")=overtime
			rs("isweek")=0
	end if
	rs.update
	call closeRs(rs)
	response.Redirect("ot_add.asp")
	end sub
	sub dell()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs.delete
	rs.update
	call closeRs(rs)
	response.Redirect("ot_add.asp")
	end sub
	sub is_ok()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("usermanager")=session("UserName")
	rs("is_ok")=Trim(Request.QueryString("is_ok"))
	rs.update
	call closeRs(rs)
	response.Redirect("ot_list.asp")
	end sub
	sub manager_ok()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("manager_ok")=0
	rs("usermanager")=session("username")
	rs("verify_time")=now()
	rs.update
	call closeRs(rs)
	response.Redirect("list_ot.asp")
	end sub
	sub manager_NG()
	call CreateRs(rs,"select * from overtime where id="&Trim(Request.QueryString("id")),1,3)
	rs("manager_ok")=2
	rs("usermanager")=session("username")
	rs("verify_time")=now()
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
				rs("verify_time")=now()	
				rs.update
			call closeRs(rs)
		end if
	next
		response.redirect("list_ot.asp")
		
end sub
call closeConn()
 %>







