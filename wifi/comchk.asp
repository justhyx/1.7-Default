<!--#include file="../connet/sconn.asp" -->
<%
chk=Trim(request.QueryString("event")) 'ע�⣬����action�ݽ�����ʱ��method="post"����Ҫд����Ȼ���ܴ�����
select case chk
case "search01"
call searchzero()
case "search02"
call searchone()
case "search03"
call searchtwo()
case "add01"
call addzero()
case "add02"
call addone()
case "add03"
call addtwo()
case "change01"
call changezero()
case "change02"
call changeone()
case "change03"
call changetwo()
case "del01"
call delzero()
case "del02"
call delone()
case "del03"
call deltwo()
case "addemail"
call addemail()
end select

sub searchzero()
checkin=Trim(request.form("checkin"))
ziduan=Trim(request.form("ziduan"))
dim sql
sql="select *from ���Թ����¼ where "+ziduan+" like "+"'%"& checkin &"%'"
response.cookies("comsql")=sql
response.Cookies("comsql").Expires=dateadd("h",24,now())
response.Redirect("computer02.asp")
end sub

sub searchone()
checkin=Trim(request.form("checkin"))
ziduan=Trim(request.form("ziduan"))
dim sql
sql="select *from ����ʹ����ϸ where "+ziduan+" like "+"'%"& checkin &"%' and ����ʹ��ʱ�� is null"
response.cookies("comsql")=sql
response.Cookies("comsql").Expires=dateadd("h",24,now())
response.Redirect("computer03.asp")
end sub

sub searchtwo()
checkin=Trim(request.form("checkin"))
ziduan=Trim(request.form("ziduan"))
dim sql
sql="select *from ����ά����¼ where "+ziduan+" like "+"'%"& checkin &"%'"
response.cookies("comsql")=sql
response.Cookies("comsql").Expires=dateadd("h",24,now())
response.Redirect("computer04.asp")
end sub

sub addzero()
mac=Trim(request.form("mac"))
buy_mon=Trim(request.form("buy_mon"))
buy_sup=Trim(request.form("buy_sup"))
buy_mold=Trim(request.form("buy_mold"))
buy_date=Trim(request.form("buy_date"))
place=Trim(request.form("place"))
randomize
ranNum=int(900*rnd)+100
if buy_mon="" or mac="" or buy_date="" then
response.Write"<script> alert('��Ϣ��д���������������Ĺ������ڣ����뵥�ۺ������ַ��'); history.back(); </script>"
response.end()
end if
call createRs(rs,"select * from ���Թ����¼ where MAC='"& mac &"'",1,3)
if rs.eof then
rs.addnew
rs("id")="BUY"&year(date())&month(date())&day(date())&ranNum
rs("MAC")=mac
rs("��������")=buy_date
rs("���뵥��")=cdbl(buy_mon)
rs("��Ӧ��")=buy_sup
rs("����")=buy_mold
rs("��λ")=1
rs("״̬")="����"
rs("λ��")=place
rs.update
else 
response.Write"<script> alert('�������ַ���ڹ����¼�У���������д��'); history.back(); </script>"
response.end()
end if
response.Redirect("computer02.asp")
call closeRS(rs)
end sub 

sub addone()
mac=Trim(request.form("mac"))
use_man=Trim(request.form("use_man"))
use_dep=Trim(request.form("use_dep"))
use_boss=Trim(request.form("use_boss"))
use_ip=Trim(request.form("use_ip"))
use_place=Trim(request.form("use_place"))
note=Trim(request.form("note"))
dbip=Trim(request.form("db_ip"))
randomize
ranNum=int(900*rnd)+100
call createRs(rs,"select * from ����ʹ����ϸ where MAC='"& mac &"' and ����ʹ��ʱ�� is null",1,3)
if rs.eof then
rs.addnew
rs("id")="USE"&year(date())&month(date())&day(date())&ranNum
rs("mac")=mac
rs("ipdress")=use_ip
rs("ʹ����")=use_man
rs("ʹ�ò���")=use_dep
rs("������")=use_boss
rs("��ʼʹ��ʱ��")=date()
rs("�칫��")=use_place
rs("˫IP")=dbip
rs("��ע")=note
rs.update()
end if
call closeRs(rs)
call createRs(rsa,"select *from ���Թ����¼ where MAC='"& mac &"'",1,3)
rsa("״̬")="ʹ����"
rsa("λ��")=use_place
rsa.update
call closeRS(rsa)
response.Redirect("computer02.asp")
end sub

sub addtwo()
fm=Trim(request.form("fm"))
thing=Trim(request.form("thing"))
reason=Trim(request.form("reason"))
sd=Trim(request.form("sd"))
fc=Trim(request.form("fc"))
ft=Trim(request.form("ft"))
zt=Trim(request.form("zt"))
it=Trim(request.form("it"))
im=Trim(request.form("im"))
bz=Trim(request.form("bz"))
xx=Trim(request.form("xx"))
call createRs(rs,"select * from ����ά����¼",1,3)
rs.addnew
randomize
ranNum=int(9000*rnd)+100
rs("id")="FIX"&year(date())&month(date())&day(date())&ranNum
rs("MAC")=fc
rs("comuser")=fm
rs("ά����")=it
rs("ά����Ŀ")=thing
rs("ά��ԭ��")=reason
rs("��ע")=bz
rs("ά������")=Cdate(ft)
rs("��������")=xx
rs("IT���")=im
rs("����ʱ��")=sd
rs("״̬")=zt
rs.update()
call closeRS(rs)
response.Redirect("computer04.asp")
end sub 

sub changezero()
buy_time=Trim(request.form("buy_time"))
buy_mon=Trim(request.form("buy_mon"))
buy_sup=Trim(request.form("buy_sup"))
buy_mold=Trim(request.form("buy_mold"))
place=Trim(request.form("place"))
mac=Trim(request.form("mac"))
zt=Trim(request.form("zt"))
id=Trim(request.form("buy_id"))
call createRs(rs,"select * from ���Թ����¼ where MAC='"& mac &"'",1,3)
if  not rs.eof then
 if rs("id")<>id then
response.Write"<script> alert('�������ַ�Ѿ���ϵͳ�����¼�У���������д��'); history.back(); </script>"
response.end()
end if
end if
call closeRS(rs)
call createRs(rs,"select * from ���Թ����¼ where id='"& id &"'",1,3)
if not rs.eof then
mac_old=rs("MAC")
if buy_time<>"" then
rs("��������")=buy_time
end if
if buy_mon<>"" then
rs("���뵥��")=Cdbl(buy_mon)
end if
rs("��Ӧ��")=buy_sup
rs("����")=buy_mold
rs("λ��")=place
rs("MAC")=mac
if zt="1" then
rs("״̬")="����"
end if
rs.update
end if
call closeRs(rs)
call createRS(rs,"select *from ����ʹ����ϸ where MAC='"& mac_old &"'",1,3)
if not rs.eof then
do while not rs.eof
rs("MAC")=mac
rs.update
rs.movenext
loop
end if
call closeRs(rs)
response.Redirect("computer02.asp")
end sub

sub changeone()
use_dep=Trim(request.form("use_dep"))
use_boss=Trim(request.form("use_boss"))
mac=Trim(request.form("mac"))
use_timeone=Trim(request.form("use_timeone"))
use_timetwo=Trim(request.form("use_timetwo"))
buy_id=Trim(request.form("buy_id"))
use_id=Trim(request.form("use_id"))
db_ip=Trim(request.form("db_ip"))
ip=Trim(request.form("ipdress"))
place=Trim(request.form("use_place"))
note=Trim(request.form("note"))
call createRs(rsm,"select * from ���Թ����¼ where MAC='"& mac &"'",1,3)
if not rsm.eof then
    if rsm("״̬")="����" or rsm("id")=buy_id then
        call createRs(rsa,"select * from ����ʹ����ϸ where id='"& use_id &"'",1,3)
        if not rsa.eof then
            rsa("ʹ�ò���")=use_dep
            rsa("������")=use_boss
            rsa("MAC")=mac
            rsa("��ʼʹ��ʱ��")=use_timeone
			if use_timetwo<>"" then
            rsa("����ʹ��ʱ��")=use_timetwo
			end if
			rsa("ipdress")=ip
			rsa("˫IP")=db_ip
			rsa("�칫��")=place
			rsa("��ע")=note
            rsa.update
        end if
        call closeRS(rsa)
        response.Redirect("computer03.asp")
    else
        response.Write"<script> alert('��Ǵ������ַ�ĵ�������ʹ�ã�������ѡ��'); history.back(); </script>"
        response.end()
    end if
else
    response.Write"<script> alert('ϵͳ���޴˵��ԣ�������ѡ��'); history.back(); </script>"
    response.end()
end if
call closeRs(rsm)
end sub

sub changetwo()
fm=Trim(request.form("fm"))
thing=Trim(request.form("thing"))
reason=Trim(request.form("reason"))
sd=Trim(request.form("sd"))
fc=Trim(request.form("fc"))
ft=Trim(request.form("ft"))
zt=Trim(request.form("zt"))
it=Trim(request.form("it"))
im=Trim(request.form("im"))
bz=Trim(request.form("bz"))
xx=Trim(request.form("xx"))
id=Trim(request.form("id"))
call createRs(rs,"select * from ����ά����¼ where id='"& id &"'",1,3)
if not rs.eof then
rs("MAC")=fc
rs("comuser")=fm
rs("ά����")=it
rs("ά����Ŀ")=thing
rs("ά��ԭ��")=reason
rs("��ע")=bz
rs("ά������")=Cdate(ft)
rs("��������")=xx
rs("IT���")=im
rs("����ʱ��")=sd
rs("״̬")=zt
rs.update()
end if
call closeRS(rs)
response.Redirect("computer04.asp")
end sub

sub delzero()
id=Trim(request.QueryString("id"))
call createRS(rs,"select * from ���Թ����¼ where id='"& id &"'",1,3)
if not rs.eof then
rs.delete
rs.update
end if
call closeRs(rs)
response.Redirect("computer02.asp")
end sub

sub delone()
id=Trim(request.QueryString("id"))
call createRS(rs,"select * from ����ʹ����ϸ where id='"& id &"'",1,3)
if not rs.eof then
mac=rs("MAC")
rs("����ʹ��ʱ��")=date()
rs.update
call createRs(rsm,"select * from ���Թ����¼ where MAC='"& mac &"'",1,3)
if not rsm.eof then
rsm("״̬")="����"
rsm("λ��")=rs("�칫��")
rsm.update
end if
call closeRs(rsm)
end if
call closeRs(rs)
response.Redirect("computer03.asp")
end sub

sub deltwo()
id=Trim(request.QueryString("id"))
call createRS(rs,"select * from ����ά����¼ where id='"& id &"'",1,3)
if not rs.eof then
rs.delete
rs.update
end if
call closeRs(rs)
response.Redirect("computer04.asp")
end sub

sub addemail()
no=Trim(request.form("no"))
eamil=Trim(request.form("email"))
call createRs(rs,"select *from HUDSON_User where UserID='"& no &"'",1,3)
if not rs.eof then
rs("email")=eamil
rs.update
else
response.Write"<script> alert('ϵͳ���޴��û���������ѡ��'); history.back(); </script>"
response.end()
end if
call closeRs(rs)
response.Write"<script> alert('��ӳɹ�'); window.location.href = 'eamiladd.asp';</script>"
end sub

call closeConn()
%>






