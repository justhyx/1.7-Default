<!--#include file="../connet/conn.asp" -->
<% 	
	'******************************
	'提交用车申请
	'******************************	 
	action=Trim(Request.QueryString("action"))
	select case action
		case "add"
			call add()
		case "modify"
			call modify()
		case "arrange"
			call arrange()
		case "do_use"
			call do_use()
	end select
sub add()
	dim usetime,Cusetime,Nusetime
	if request.Form("cartype")="0" then
		response.Write("<script> alert('请选择用车类型'); history.back(); </script>")
		response.End()
	end if
	call CreateRs(rs,"select * from UseCar",1,3)
	rs.addnew
	if Trim(request.Form("Urgency"))<>"" then
		rs("Urgencyfloy")=2
		rs("arrange")=1
	else
		rs("Urgencyfloy")=1
		rs("arrange")=1
	end if
	if Trim(Request.Form("UsePeople"))="" or Trim(Request.Form("Reason"))="" then
		response.Write"<script>alert('请注意：使用人和事由必须填写'); history.back();</script>"
		response.End()
	end if
	usetime=Trim(Request.Form("StartTime"))&" "&Trim(Request.Form("hh"))&":"&Trim(Request.Form("mm"))
	Nusetime=now()		
	if datediff("n",cdate(cdate(Nusetime)),cdate(usetime))<180 and Trim(Request.Form("Urgency"))="" then
	response.Write"<script>alert('请注意：没有提前3小时申请车需要填写紧急用车理由'); history.back();</script>"
	response.End()
	end if
	'if  Trim(Request.Form("Urgency"))="" and( weekday(date)=0 or weekday(date)=7) then
'	response.Write"<script>alert('请注意：周末申请车需要填写紧急用车理由'); history.back();<script>"
'	response.End()
'	end if
	rs("CarDepartment")=Trim(Request.Form("CarDepartment"))
	rs("UsePeople")=Trim(Request.Form("UsePeople"))
	rs("StartTime")=Trim(Request.Form("StartTime"))
	rs("OverTime")=Trim(Request.Form("OverTime"))
	rs("Destination")=Trim(Request.Form("Destination"))
	rs("UseTime")=Trim(Request.Form("UseTime"))
	rs("Reason")=Trim(Request.Form("Reason"))
	rs("Urgency")=Trim(Request.Form("Urgency"))
	rs("Improve")=Trim(Request.Form("Improve"))
	rs("Agree")=2
	rs("UrgencyAgree")=2
	rs("Addtime")=now()
	rs("AboutTime")=usetime
	rs("pl")=Trim(Request.Form("pl"))
	rs("delfolg")=0
	rs("over_l")=session("over_l")
	rs("cartype")=request.Form("cartype")
	rs("peopleqty")=request.Form("peopleqty")
	rs.update
	call closeRs(rs)
	'******************************
	'操作日志添加
	'******************************
	call CreateRs(rs,"select * from journal",1,3)
		rs.addnew
		rs("username")=session("UserName")
		rs("usetime")=now()
		rs("usedo")="添加用车申请"
		rs.update
	call closeRs(rs)
'	select case session("over_l")
'		case "3"			
'			Email = "yu_pang@sz-hudson.com"         '收件人邮箱Email
'		case"2"
'			Email = "hua_zhang@sz-hudson.com"
'		case"7"
'			Email = "yun_zhang@sz-hudson.com"
'		case"5"
'			Email = "yanjun_yu@sz-hudson.com"
'		case"9"
'			Email = "mei_wu@sz-hudson.com"
'		case"4"
'			Email = "sunnydai@sz-hudson.com"
'		case"8"
'			Email = "tina_hu@sz-hudson.com"
'		case"6"
'			Email = "zhongmei_chen@sz-hudson.com"
'		case"1"
'			Email = "ouyang_cnc@sz-hudson.com"
'		case"0"
'			Email = "jianchun_lin@sz-hudson.com"
'	end select	
'	Set jmail = Server.CreateObject("JMAIL.Message")        '建立发送邮件的对象
'	jmail.silent = false                          		    '屏蔽例外错误，返回FALSE跟TRUE两值
'	jmail.logging = true                          		    '启用邮件日志
'	jmail.Charset = "GB2312"                      	    	'邮件的文字编码为国标
'	jmail.ContentType = "text/html"              		    '邮件的格式为HTML格式
'	jmail.AddRecipient Email                     		    '邮件收件人的地址
'	jmail.From = "systemmail@sz-hudson.com"                 '发件人的E-MAIL地址
'	jmail.MailServerUserName = "systemmail@sz-hudson.com"   '登录邮件服务器所需的用户名
'	jmail.MailServerPassword = "Hh123456h"                    '登录邮件服务器所需的密码
'	jmail.Subject = "用车申请"                 	            '邮件的标题 
'	jmail.Body = "本部有用车申请审核。"                     '邮件的内容
'	jmail.Priority = 1                         		        '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
'	jmail.Send("smtp.exmail.qq.com")          	            '执行邮件发送（通过邮件服务器地址）。
'	jmail.Close()                       		            '关闭对象
'	response.Redirect("UseList.asp")
end sub
sub modify()
	'******************************
	'审核用车申请
	'******************************
	call CreateRs(rs,"select * from UseCar where id="&Trim(Request.QueryString("id")),1,3)
	rs("Agree")=Trim(Request.Form("Agree"))
	rs("UrgencyAgree")=Trim(Request.Form("UrgencyAgree"))
	rs("Improve")=Trim(Request.Form("Improve"))
	'rs("Urgencyfloy")=0
	rs("ArrangeTime")=now()
	rs("Agreepeople")=session("UserName")
	rs.update
	call closeRs(rs)
	'******************************
	'操作日志添加
	'******************************
	call CreateRs(rs,"select * from journal",1,3)
		rs.addnew
		rs("username")=session("UserName")
		rs("usetime")=now()
		rs("usedo")="审核了用车申请"
		rs.update
	call closeRs(rs)
	response.Redirect("UseList.asp")
end sub
	'******************************
	'确定派车
	'******************************
sub arrange()
	call CreateRs(rs,"select * from UseCar where id="&Trim(Request.QueryString("id")),1,3)
	rs("delfolg")=int(Trim(Request.Form("delfolg")))
	rs("arrange")=0  
	rs("addarrangetime")=now()
	rs("CarNumber")=Trim(Request.Form("CarNumber"))
	rs("CarPeople")=Trim(Request.Form("CarPeople"))
	rs("note")=Trim(Request.Form("note"))
	rs.update
	call closeRs(rs)
	'******************************
	'操作日志添加
	'******************************
	if request.Form("delfolg")=2 then
	call CreateRs(rs,"select * from journal",1,3)
		rs.addnew
		rs("username")=session("UserName")
		rs("usetime")=now()
		rs("usedo")="删除了用车申请单"
		rs.update
	call closeRs(rs)
	end if
	response.Redirect("personnel_list.asp")
end sub
sub do_use()
call CreateRs(rs,"select * from useCar where id="&Trim(Request.QueryString("id")),1,3)
rs("serverp")=Trim(Request.Form("serverp"))
rs("carstyle")=Trim(Request.Form("carstyle"))
rs("do_use")=Trim(Request.Form("do_use"))
rs("fhbc")=Trim(Request.Form("fhbc"))
rs("gslf")=Trim(Request.Form("gslf"))
rs("ddxh")=Trim(Request.Form("ddxh"))
rs("hjyf")=Trim(Request.Form("hjyf"))
rs.update
call closeRs(rs)
response.Redirect("personnel_list.asp")
end sub
	call closeConn()
 %>











