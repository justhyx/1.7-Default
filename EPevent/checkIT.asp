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
response.Write"<script> alert('�豸�ź��������ݱ�����д'); history.back(); </script>"
response.end()
end if
if itel="" or detail="" or user="" then
response.Write"<script> alert('��Ϣ��д���������������������ˣ��������������������'); history.back(); </script>"
response.end()
end if
if phone="" and tele="" then
response.Write"<script> alert('��Ϣ��д���������ֻ�/�̺ź����ߵ绰������һ�'); history.back(); </script>"
response.end()
end if
call CreateRs(rs,"select * from Epevent",1,3)
rs.addnew
randomize
ranNum=int(90000*rnd)+10000
rs("id")="FIX"&year(date())&month(date())&day(date())&ranNum
rs("comuser")=user
rs("ά����Ŀ")=itel
rs("��������")=detail
rs("phone")=phone
rs("tele")=tele
rs("����ʱ��")=now()
rs("״̬")="δ����"
rs("nq")=nq
rs("jj")=jj
rs("equip_no")=Trim(Request.Form("equip_no"))
rs("ylnr")=Trim(Request.Form("ylnr"))
rs("chk")="δ��"
rs.update
call closeRs(rs)
response.Write"<script> alert('�ύ�ɹ������������ᷢ���������������У�'); window.location.href = 'eventlook.asp';</script>" 
end sub

sub IT_query()
id=Trim(Request.form("whid"))
call CreateRs(rsa,"select * from Epevent where id='"& id &"'",1,3)
rsa("ά������")=now
rsa("�������")=Trim(Request.Form("jy"))
rsa("״̬")="�Ѵ���"
rsa("ά����")=Trim(Request.Form("whr"))
rsa.update
comuser=rsa("comuser")

call CreateRs(rs,"select * from HUDSON_User where UserID='"& comuser &"' and email is not null",1,3)
if not rs.eof then
email=rs("email")
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
	jmail.Subject = "�豸ά������"                 	            '�ʼ��ı��� 
	jmail.Body = "���� "+cstr(rsa("����ʱ��"))+" �ύ���豸ά�޸��������Ѿ�������ϡ�ά�޸�������Ϊ "+cstr(rsa("��������"))+",�������: "+cstr(rsa("�������"))+",����ʱ��Ϊ: "+cstr(rsa("ά������"))+"��"                     '�ʼ�������
	jmail.Priority = 1                         		        '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
	jmail.Send("smtp.exmail.qq.com")          	            'ִ���ʼ����ͣ�ͨ���ʼ���������ַ����
	jmail.Close()      
	call closeRs(rsa)
end if
response.Redirect("eventlook.asp")
end sub

sub IT_result

end sub
call closeConn()
 %>







