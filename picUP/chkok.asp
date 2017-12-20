<!--#include file="../connet/conn.asp" -->
<% 
dim ID,PicName,Newfolg,maseg
maseg=Trim(Request.QueryString("maseg"))
ID=Trim(Request.QueryString("id"))
if maseg="ok" then                         '确认新图
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)
	rs("okfolg")=0
	rs("okpeople")=session("UserName")
	rs("okTime")=now()
	Newfolg=rs("newfolg")
	PicName=UCase(rs("Name"))
	rs.update
call closeRs(rs)
'************************************
'          自动判断图档类型(新/旧)
'************************************
call CreateRs(rs,"select * from PicFiles where [name]='"&PicName&"'",1,3)  '相同图番全部变为旧图档
	dim i
	i=0
	do while not rs.eof
	rs("newfolg")=0
	rs.update
	i=i+1
	rs.movenext
	loop	
call closeRs(rs)
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)   '确认当前条变为新图档
	rs("newfolg")=1
	mesnumber=rs("name")
	mesbanben=rs("suffix")
	mestime=rs("uptime")
	mesmanager=rs("okpeople")
	rs.update
	call closeRs(rs)
Email = "picflies@sz-hudson.com"                      '收件人信箱
Set jmail = Server.CreateObject("JMAIL.Message")    '建立发送邮件的对象
jmail.silent = false                          		'屏蔽例外错误，返回FALSE跟TRUE两值
jmail.logging = true                          		'启用邮件日志
jmail.Charset = "GB2312"                      		'邮件的文字编码为国标
jmail.ContentType = "text/html"              		 '邮件的格式为HTML格式
jmail.AddRecipient Email                     		 '邮件收件人的地址
jmail.From = "systemmail@sz-hudson.com"                 '发件人的E-MAIL地址
jmail.MailServerUserName = "systemmail@sz-hudson.com"   '登录邮件服务器所需的用户名
jmail.MailServerPassword = "Hh123456h"             '登录邮件服务器所需的密码
jmail.Subject = "新图档或者更新图档下载"                 		  '邮件的标题 
jmail.Body = "有图档更新，请<a href='http://192.168.1.7/PicFiles.asp'>点击进入主页查看下载</a>或者在地址栏键入http://192.168.1.7/PicFiles.asp。图档内容--更新图番："&mesnumber&"  更新版本："&mesbanben&"  图档时间："&mestime&"  图档担当："&mesmanager  '邮件的内容
jmail.Priority = 1                         		   '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
jmail.Send("smtp.exmail.qq.com")          	        '执行邮件发送（通过邮件服务器地址）。
jmail.Close()                          		        '关闭对象
response.Redirect("piclist.asp")
'图档错误处理
elseif maseg="no" then
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)
nameid=rs("Name")&rs("Suffix")
namepeople=session("UserName")
rs.delete
rs.update
call closeRs(rs)
Email = "shuangshuang_wu@sz-hudson.com"                      '收件人信箱
Set jmail = Server.CreateObject("JMAIL.Message")    '建立发送邮件的对象
jmail.silent = false                          		'屏蔽例外错误，返回FALSE跟TRUE两值
jmail.logging = true                          		'启用邮件日志
jmail.Charset = "GB2312"                      		'邮件的文字编码为国标
jmail.ContentType = "text/html"              		 '邮件的格式为HTML格式
jmail.AddRecipient Email                     		 '邮件收件人的地址
jmail.From = "systemmail@sz-hudson.com"                 '发件人的E-MAIL地址
jmail.MailServerUserName = "systemmail@sz-hudson.com"   '登录邮件服务器所需的用户名
jmail.MailServerPassword = "hudson"             '登录邮件服务器所需的密码
jmail.Subject = "有图档错误"                 		  '邮件的标题 
jmail.Body = "有图档上传错误，请与担当者确认。错误图番"&nameid&"担当者"&namepeople   '邮件的内容
jmail.Priority = 1                         		   '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
jmail.Send("mail.sz-hudson.com")          	        '执行邮件发送（通过邮件服务器地址）。
jmail.Close()                          		        '关闭对象
response.Redirect("piclist.asp")          
elseif maseg="del" then                             '删除上传错误图档
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)  
rs.delete
rs.update
call closeRs(rs)
response.Redirect("Upload_Photo.asp") 
end if
call closeConn()
 %>

