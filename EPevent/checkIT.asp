<!--#include file="../connet/conn.asp" -->
<% 
dim action
action=Trim(Request.QueryString("action"))
select case action
	case "add"
		call IT_add()
	case "IT_query"
		call IT_query()
	case "IT_result"
		call IT_result()
end select
sub IT_add()
user=Trim(request.form("name"))
itel=Trim(request.form("IT_Tel"))
nq=Trim(request.form("nq"))
detail=Trim(request.form("itdetail"))
phone=Trim(request.form("phone"))
tele=Trim(request.form("tele"))
jj=Trim(request.form("jj"))
equip_no=Trim(Request.Form("equip_no"))
ylnr=Trim(Request.Form("ylnr"))
if equip_no="" or ylnr="" then
response.Write"<script> alert('设备号和依赖内容必须填写'); history.back(); </script>"
response.end()
end if
if itel="" or detail="" or user="" then
response.Write"<script> alert('信息填写不完整，请检查您的依赖人，依赖事项和现象描述！'); history.back(); </script>"
response.end()
end if
if phone="" and tele="" then
response.Write"<script> alert('信息填写不完整，手机/短号和内线电话必须填一项。'); history.back(); </script>"
response.end()
end if
call CreateRs(rs,"select * from Epevent",1,3)
rs.addnew
randomize
ranNum=int(90000*rnd)+10000
rs("id")="FIX"&year(date())&month(date())&day(date())&ranNum
rs("comuser")=user
rs("维护项目")=itel
rs("现象描述")=detail
rs("phone")=phone
rs("tele")=tele
rs("依赖时间")=now()
rs("状态")="未处理"
rs("nq")=nq
rs("jj")=jj
rs("equip_no")=Trim(Request.Form("equip_no"))
rs("ylnr")=Trim(Request.Form("ylnr"))
rs("chk")="未审"
rs.update
call closeRs(rs)
response.Write"<script> alert('提交成功，处理结果将会发送至依赖人邮箱中！'); window.location.href = 'eventlook.asp';</script>" 
end sub

sub IT_query()
id=Trim(Request.form("whid"))
call CreateRs(rsa,"select * from Epevent where id='"& id &"'",1,3)
rsa("维护日期")=now
rsa("处置情况")=Trim(Request.Form("jy"))
rsa("状态")="已处理"
rsa("维护人")=Trim(Request.Form("whr"))
rsa.update
comuser=rsa("comuser")

call CreateRs(rs,"select * from HUDSON_User where UserID='"& comuser &"' and email is not null",1,3)
if not rs.eof then
email=rs("email")
call closeRs(rs)
	Set jmail = Server.CreateObject("JMAIL.Message")        '建立发送邮件的对象
	jmail.silent = false                          		    '屏蔽例外错误，返回FALSE跟TRUE两值
	jmail.logging = true                          		    '启用邮件日志
	jmail.Charset = "GB2312"                      	    	'邮件的文字编码为国标
	jmail.ContentType = "text/html"              		    '邮件的格式为HTML格式
	jmail.AddRecipient email                     		    '邮件收件人的地址
	jmail.From = "systemmail@sz-hudson.com"                 '发件人的E-MAIL地址
	jmail.MailServerUserName = "systemmail@sz-hudson.com"   '登录邮件服务器所需的用户名
	jmail.MailServerPassword = "hudson1"                    '登录邮件服务器所需的密码
	jmail.Subject = "设备维护依赖"                 	            '邮件的标题 
	jmail.Body = "您于 "+cstr(rsa("依赖时间"))+" 提交的设备维修改善依赖已经处理完毕。维修改善描述为 "+cstr(rsa("现象描述"))+",处理情况: "+cstr(rsa("处置情况"))+",处理时间为: "+cstr(rsa("维护日期"))+"。"                     '邮件的内容
	jmail.Priority = 1                         		        '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
	jmail.Send("smtp.exmail.qq.com")          	            '执行邮件发送（通过邮件服务器地址）。
	jmail.Close()      
	call closeRs(rsa)
end if
response.Redirect("eventlook.asp")
end sub

sub IT_result

end sub
call closeConn()
 %>







