<!--#include file="../connet/conn.asp" -->
<% 
dim ID,PicName,Newfolg,maseg
maseg=Trim(Request.QueryString("maseg"))
ID=Trim(Request.QueryString("id"))
if maseg="ok" then                         'ȷ����ͼ
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)
	rs("okfolg")=0
	rs("okpeople")=session("UserName")
	rs("okTime")=now()
	Newfolg=rs("newfolg")
	PicName=UCase(rs("Name"))
	rs.update
call closeRs(rs)
'************************************
'          �Զ��ж�ͼ������(��/��)
'************************************
call CreateRs(rs,"select * from PicFiles where [name]='"&PicName&"'",1,3)  '��ͬͼ��ȫ����Ϊ��ͼ��
	dim i
	i=0
	do while not rs.eof
	rs("newfolg")=0
	rs.update
	i=i+1
	rs.movenext
	loop	
call closeRs(rs)
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)   'ȷ�ϵ�ǰ����Ϊ��ͼ��
	rs("newfolg")=1
	mesnumber=rs("name")
	mesbanben=rs("suffix")
	mestime=rs("uptime")
	mesmanager=rs("okpeople")
	rs.update
	call closeRs(rs)
Email = "picflies@sz-hudson.com"                      '�ռ�������
Set jmail = Server.CreateObject("JMAIL.Message")    '���������ʼ��Ķ���
jmail.silent = false                          		'����������󣬷���FALSE��TRUE��ֵ
jmail.logging = true                          		'�����ʼ���־
jmail.Charset = "GB2312"                      		'�ʼ������ֱ���Ϊ����
jmail.ContentType = "text/html"              		 '�ʼ��ĸ�ʽΪHTML��ʽ
jmail.AddRecipient Email                     		 '�ʼ��ռ��˵ĵ�ַ
jmail.From = "systemmail@sz-hudson.com"                 '�����˵�E-MAIL��ַ
jmail.MailServerUserName = "systemmail@sz-hudson.com"   '��¼�ʼ�������������û���
jmail.MailServerPassword = "Hh123456h"             '��¼�ʼ����������������
jmail.Subject = "��ͼ�����߸���ͼ������"                 		  '�ʼ��ı��� 
jmail.Body = "��ͼ�����£���<a href='http://192.168.1.7/PicFiles.asp'>���������ҳ�鿴����</a>�����ڵ�ַ������http://192.168.1.7/PicFiles.asp��ͼ������--����ͼ����"&mesnumber&"  ���°汾��"&mesbanben&"  ͼ��ʱ�䣺"&mestime&"  ͼ��������"&mesmanager  '�ʼ�������
jmail.Priority = 1                         		   '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
jmail.Send("smtp.exmail.qq.com")          	        'ִ���ʼ����ͣ�ͨ���ʼ���������ַ����
jmail.Close()                          		        '�رն���
response.Redirect("piclist.asp")
'ͼ��������
elseif maseg="no" then
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)
nameid=rs("Name")&rs("Suffix")
namepeople=session("UserName")
rs.delete
rs.update
call closeRs(rs)
Email = "shuangshuang_wu@sz-hudson.com"                      '�ռ�������
Set jmail = Server.CreateObject("JMAIL.Message")    '���������ʼ��Ķ���
jmail.silent = false                          		'����������󣬷���FALSE��TRUE��ֵ
jmail.logging = true                          		'�����ʼ���־
jmail.Charset = "GB2312"                      		'�ʼ������ֱ���Ϊ����
jmail.ContentType = "text/html"              		 '�ʼ��ĸ�ʽΪHTML��ʽ
jmail.AddRecipient Email                     		 '�ʼ��ռ��˵ĵ�ַ
jmail.From = "systemmail@sz-hudson.com"                 '�����˵�E-MAIL��ַ
jmail.MailServerUserName = "systemmail@sz-hudson.com"   '��¼�ʼ�������������û���
jmail.MailServerPassword = "hudson"             '��¼�ʼ����������������
jmail.Subject = "��ͼ������"                 		  '�ʼ��ı��� 
jmail.Body = "��ͼ���ϴ��������뵣����ȷ�ϡ�����ͼ��"&nameid&"������"&namepeople   '�ʼ�������
jmail.Priority = 1                         		   '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
jmail.Send("mail.sz-hudson.com")          	        'ִ���ʼ����ͣ�ͨ���ʼ���������ַ����
jmail.Close()                          		        '�رն���
response.Redirect("piclist.asp")          
elseif maseg="del" then                             'ɾ���ϴ�����ͼ��
call CreateRs(rs,"select * from PicFiles where id="&ID,1,3)  
rs.delete
rs.update
call closeRs(rs)
response.Redirect("Upload_Photo.asp") 
end if
call closeConn()
 %>

