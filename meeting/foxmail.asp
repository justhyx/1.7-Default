<%   
Email = "hr01@sz-hudson.com"
Set jmail = Server.CreateObject("JMAIL.Message")    '���������ʼ��Ķ���
jmail.silent = false                          		'����������󣬷���FALSE��TRUE��ֵ
jmail.logging = true                          		'�����ʼ���־
jmail.Charset = "GB2312"                      		'�ʼ������ֱ���Ϊ����
jmail.ContentType = "text/html"              		 '�ʼ��ĸ�ʽΪHTML��ʽ
jmail.AddRecipient Email                     		 '�ʼ��ռ��˵ĵ�ַ
jmail.From = "systemmail@sz-hudson.com"                 '�����˵�E-MAIL��ַ
jmail.MailServerUserName = "systemmail@sz-hudson.com"   '��¼�ʼ�������������û���
jmail.MailServerPassword = "Hh123456h"             '��¼�ʼ����������������
jmail.Subject = "������ʹ��"                 		  '�ʼ��ı��� 
jmail.Body = "<a href='http://192.168.1.7/login.asp'>�����¼�鿴</a>��"   '�ʼ�������
jmail.Priority = 1                         		   '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
jmail.Send("smtp.exmail.qq.com")         	        'ִ���ʼ����ͣ�ͨ���ʼ���������ַ�������޸ĳ�����ʼ�������SMTP��ַ
jmail.Close()                        		     '�رն���
response.Redirect("meetinglist.asp")
%>
