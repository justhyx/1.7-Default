<%
select case session("Department")
	case "��Ʒ������"                         
		Email = "yanjun_yu@sz-hudson.com"         '�ռ�������Email
	case"������"
		Email = "hua_zhang@sz-hudson.com"
	case"���������"
		Email = "yun_zhang@sz-hudson.com"
	case"���ο�"
		Email = "zhen_xie@sz-hudson.com"
	case"ע�ܼ���"
		Email = "mei_wu@sz-hudson.com"
	case"�����"
		Email = "sunnydai@sz-hudson.com"
	case"Ӫҵ��","Ӫҵ�ɹ���"
		Email = "tina_hu@sz-hudson.com"
	case"���������"
		Email = "zhongmei_chen@sz-hudson.com"
	case"������"
		Email = "ouyang_cnc@sz-hudson.com"
	case"�ֿ�����"
		Email = "jianchun_lin@sz-hudson.com"
end select

Set jmail = Server.CreateObject("JMAIL.Message")    '���������ʼ��Ķ���
jmail.silent = false                          		'����������󣬷���FALSE��TRUE��ֵ
jmail.logging = true                          		'�����ʼ���־
jmail.Charset = "GB2312"                      		'�ʼ������ֱ���Ϊ����
jmail.ContentType = "text/html"              		 '�ʼ��ĸ�ʽΪHTML��ʽ
jmail.AddRecipient Email                     		 '�ʼ��ռ��˵ĵ�ַ
jmail.From = "ke_wan@sz-hudson.com"                 '�����˵�E-MAIL��ַ
jmail.MailServerUserName = "ke_wan@sz-hudson.com"   '��¼�ʼ�������������û���
jmail.MailServerPassword = "wankejie28"             '��¼�ʼ����������������
jmail.Subject = "�ó�����"                 		  '�ʼ��ı��� 
jmail.Body = "�������ó����룬��<a href='http://192.168.1.7/UseCar/UseList.asp'>���������˲鿴</a>��"   '�ʼ�������
'jmail.appendtext("<a href='http://192.168.1.7/UseCar/UseList.asp' >����</a>")
jmail.Priority = 1                         		   '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
jmail.Send("mail.sz-hudson.com")          	        'ִ���ʼ����ͣ�ͨ���ʼ���������ַ�������޸ĳ�����ʼ�������SMTP��ַ
jmail.Close()                            		     '�رն���

response.Redirect("UseList.asp")
%>
