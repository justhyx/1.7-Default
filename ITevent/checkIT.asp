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
itevent=Trim(request.form("itevent"))
detail=Trim(request.form("itdetail"))
phone=Trim(request.form("phone"))
tele=Trim(request.form("tele"))
if itel="" or detail="" or user="" then
response.Write"<script> alert('��Ϣ��д���������������������ˣ��������������������'); history.back(); </script>"
response.end()
end if
if phone="" and tele="" then
response.Write"<script> alert('��Ϣ��д���������ֻ�/�̺ź����ߵ绰������һ�'); history.back(); </script>"
response.end()
end if
call CreateRs(rs,"select * from ����ά����¼",1,3)
rs.addnew
randomize
ranNum=int(9000*rnd)+100
rs("id")="FIX"&year(date())&month(date())&day(date())&ranNum
rs("comuser")=user
rs("ά����Ŀ")=itel
rs("ά��ԭ��")=itevent
rs("��������")=detail
rs("phone")=phone
rs("tele")=tele
rs("����ʱ��")=now()
rs("״̬")="δ����"
rs.update
call closeRs(rs)
call createRS(rs,"select *from HUDSON_User where UserID='"& user &"'",1,1)
if not rs.eof then
username=cstr(rs("username"))
else
username="��������"
end if
call closeRs(rs)
 
	Set jmail = Server.CreateObject("JMAIL.Message")        '���������ʼ��Ķ���
	jmail.silent = false                          		    '����������󣬷���FALSE��TRUE��ֵ
	jmail.logging = true                          		    '�����ʼ���־
	jmail.Charset = "GB2312"                      	    	'�ʼ������ֱ���Ϊ����
	jmail.ContentType = "text/html"              		    '�ʼ��ĸ�ʽΪHTML��ʽ
	jmail.AddRecipient "chulong_wang@sz-hudson.com"                    		    '�ʼ��ռ��˵ĵ�ַ
	jmail.From = "systemmail@sz-hudson.com"                 '�����˵�E-MAIL��ַ
	jmail.MailServerUserName = "systemmail@sz-hudson.com"   '��¼�ʼ�������������û���
	jmail.MailServerPassword = "hudson1"                    '��¼�ʼ����������������
	jmail.Subject = "IT��������"                 	            '�ʼ��ı��� 
	jmail.Body = "�û� "+username+" �� "+cstr(now())+" ����������һ��IT�������뵥���������"+itel+"�� ��������:"+detail+" "'�ʼ�������
	jmail.Priority = 1                         		        '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
	jmail.Send("smtp.exmail.qq.com")          	            'ִ���ʼ����ͣ�ͨ���ʼ���������ַ����
	jmail.Close()   
response.Write"<script> alert('�ύ�ɹ������������ᷢ���������������У�'); window.location.href = 'eventlook.asp';</script>"
end sub

sub IT_query()
id=Trim(Request.form("whid"))
ipdress=Trim(Request.form("ipdress"))
itmind=Trim(request.form("itmind"))
call CreateRs(rs,"select * from ����ʹ����ϸ where ipdress='"& ipdress &"'",1,3)
if not rs.eof then
mac=rs("MAC")
else
mac=""
end if
call closeRs(rs)
call CreateRs(rsa,"select * from ����ά����¼ where id='"& id &"'",1,3)
rsa("MAC")=mac
rsa("ά������")=now
rsa("IT���")=itmind
rsa("״̬")="�Ѵ���"
rsa("ά����")=session("username")
rsa.update
comuser=rsa("comuser")

call CreateRs(rs,"select * from HUDSON_User where UserID='"& comuser &"'",1,3)
if not rs.eof then
email=rs("email")
else
email="chulong_wang@sz-hudson.com"
end if
call closeRs(rs)
	Set jmail = Server.CreateObject("JMAIL.Message")        '���������ʼ��Ķ���
	jmail.silent = false                          		    '����������󣬷���FALSE��TRUE��ֵ
	jmail.logging = true                          		    '�����ʼ���־
	jmail.Charset = "GB2312"                      	    	'�ʼ������ֱ���Ϊ����
	jmail.ContentType = "text/html"              		    '�ʼ��ĸ�ʽΪHTML��ʽ
	jmail.AddRecipient email                     		    '�ʼ��ռ��˵ĵ�ַ
	jmail.From = "systemmail@sz-hudson.com"                 '�����˵�E-MAIL��ַ
	jmail.MailServerUserName = "systemmail@sz-hudson.com"   '��¼�ʼ�������������û���
	jmail.MailServerPassword = "hudson1"                    '��¼�ʼ����������������
	jmail.Subject = "IT��������"                 	            '�ʼ��ı��� 
	jmail.Body = "���� "+cstr(rsa("����ʱ��"))+" �ύ��IT�������뵥�Ѿ�������ϡ���������Ϊ: "+cstr(rsa("ά����Ŀ"))+"����������Ϊ "+cstr(rsa("��������"))+",IT���Ϊ: "+cstr(rsa("IT���"))+",IT����ʱ��Ϊ: "+cstr(rsa("ά������"))+"��"                     '�ʼ�������
	jmail.Priority = 1                         		        '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
	jmail.Send("smtp.exmail.qq.com")          	            'ִ���ʼ����ͣ�ͨ���ʼ���������ַ����
	jmail.Close()      
	call closeRs(rsa)
response.Redirect("eventlook.asp")
end sub

sub IT_result

end sub
call closeConn()
 %>







