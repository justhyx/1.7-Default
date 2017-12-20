<%
select case session("Department")
	case "部品技术课"                         
		Email = "yanjun_yu@sz-hudson.com"         '收件人邮箱Email
	case"技术课"
		Email = "hua_zhang@sz-hudson.com"
	case"生产管理课"
		Email = "yun_zhang@sz-hudson.com"
	case"成形课"
		Email = "zhen_xie@sz-hudson.com"
	case"注塑检查课"
		Email = "mei_wu@sz-hudson.com"
	case"财务课"
		Email = "sunnydai@sz-hudson.com"
	case"营业课","营业采购课"
		Email = "tina_hu@sz-hudson.com"
	case"人事总务课"
		Email = "zhongmei_chen@sz-hudson.com"
	case"生产课"
		Email = "ouyang_cnc@sz-hudson.com"
	case"仓库管理课"
		Email = "jianchun_lin@sz-hudson.com"
end select

Set jmail = Server.CreateObject("JMAIL.Message")    '建立发送邮件的对象
jmail.silent = false                          		'屏蔽例外错误，返回FALSE跟TRUE两值
jmail.logging = true                          		'启用邮件日志
jmail.Charset = "GB2312"                      		'邮件的文字编码为国标
jmail.ContentType = "text/html"              		 '邮件的格式为HTML格式
jmail.AddRecipient Email                     		 '邮件收件人的地址
jmail.From = "ke_wan@sz-hudson.com"                 '发件人的E-MAIL地址
jmail.MailServerUserName = "ke_wan@sz-hudson.com"   '登录邮件服务器所需的用户名
jmail.MailServerPassword = "wankejie28"             '登录邮件服务器所需的密码
jmail.Subject = "用车申请"                 		  '邮件的标题 
jmail.Body = "本部有用车申请，请<a href='http://192.168.1.7/UseCar/UseList.asp'>点击进入审核查看</a>。"   '邮件的内容
'jmail.appendtext("<a href='http://192.168.1.7/UseCar/UseList.asp' >请点击</a>")
jmail.Priority = 1                         		   '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
jmail.Send("mail.sz-hudson.com")          	        '执行邮件发送（通过邮件服务器地址）。请修改成你的邮件服务器SMTP地址
jmail.Close()                            		     '关闭对象

response.Redirect("UseList.asp")
%>
