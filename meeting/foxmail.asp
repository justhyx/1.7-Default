<%   
Email = "hr01@sz-hudson.com"
Set jmail = Server.CreateObject("JMAIL.Message")    '建立发送邮件的对象
jmail.silent = false                          		'屏蔽例外错误，返回FALSE跟TRUE两值
jmail.logging = true                          		'启用邮件日志
jmail.Charset = "GB2312"                      		'邮件的文字编码为国标
jmail.ContentType = "text/html"              		 '邮件的格式为HTML格式
jmail.AddRecipient Email                     		 '邮件收件人的地址
jmail.From = "systemmail@sz-hudson.com"                 '发件人的E-MAIL地址
jmail.MailServerUserName = "systemmail@sz-hudson.com"   '登录邮件服务器所需的用户名
jmail.MailServerPassword = "Hh123456h"             '登录邮件服务器所需的密码
jmail.Subject = "会议室使用"                 		  '邮件的标题 
jmail.Body = "<a href='http://192.168.1.7/login.asp'>点击登录查看</a>。"   '邮件的内容
jmail.Priority = 1                         		   '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
jmail.Send("smtp.exmail.qq.com")         	        '执行邮件发送（通过邮件服务器地址）。请修改成你的邮件服务器SMTP地址
jmail.Close()                        		     '关闭对象
response.Redirect("meetinglist.asp")
%>
