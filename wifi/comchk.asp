<!--#include file="../connet/sconn.asp" -->
<%
chk=Trim(request.QueryString("event")) '注意，在用action递交参数时，method="post"必须要写，不然不能传参数
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
sql="select *from 电脑购入记录 where "+ziduan+" like "+"'%"& checkin &"%'"
response.cookies("comsql")=sql
response.Cookies("comsql").Expires=dateadd("h",24,now())
response.Redirect("computer02.asp")
end sub

sub searchone()
checkin=Trim(request.form("checkin"))
ziduan=Trim(request.form("ziduan"))
dim sql
sql="select *from 电脑使用明细 where "+ziduan+" like "+"'%"& checkin &"%' and 结束使用时间 is null"
response.cookies("comsql")=sql
response.Cookies("comsql").Expires=dateadd("h",24,now())
response.Redirect("computer03.asp")
end sub

sub searchtwo()
checkin=Trim(request.form("checkin"))
ziduan=Trim(request.form("ziduan"))
dim sql
sql="select *from 电脑维护记录 where "+ziduan+" like "+"'%"& checkin &"%'"
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
response.Write"<script> alert('信息填写不完整，请检查您的购入日期，购入单价和物理地址！'); history.back(); </script>"
response.end()
end if
call createRs(rs,"select * from 电脑购入记录 where MAC='"& mac &"'",1,3)
if rs.eof then
rs.addnew
rs("id")="BUY"&year(date())&month(date())&day(date())&ranNum
rs("MAC")=mac
rs("购入日期")=buy_date
rs("购入单价")=cdbl(buy_mon)
rs("供应商")=buy_sup
rs("类型")=buy_mold
rs("单位")=1
rs("状态")="闲置"
rs("位置")=place
rs.update
else 
response.Write"<script> alert('此物理地址已在购入记录中，请重新填写！'); history.back(); </script>"
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
call createRs(rs,"select * from 电脑使用明细 where MAC='"& mac &"' and 结束使用时间 is null",1,3)
if rs.eof then
rs.addnew
rs("id")="USE"&year(date())&month(date())&day(date())&ranNum
rs("mac")=mac
rs("ipdress")=use_ip
rs("使用人")=use_man
rs("使用部门")=use_dep
rs("责任人")=use_boss
rs("开始使用时间")=date()
rs("办公室")=use_place
rs("双IP")=dbip
rs("备注")=note
rs.update()
end if
call closeRs(rs)
call createRs(rsa,"select *from 电脑购入记录 where MAC='"& mac &"'",1,3)
rsa("状态")="使用中"
rsa("位置")=use_place
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
call createRs(rs,"select * from 电脑维护记录",1,3)
rs.addnew
randomize
ranNum=int(9000*rnd)+100
rs("id")="FIX"&year(date())&month(date())&day(date())&ranNum
rs("MAC")=fc
rs("comuser")=fm
rs("维护人")=it
rs("维护项目")=thing
rs("维护原因")=reason
rs("备注")=bz
rs("维护日期")=Cdate(ft)
rs("现象描述")=xx
rs("IT意见")=im
rs("依赖时间")=sd
rs("状态")=zt
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
call createRs(rs,"select * from 电脑购入记录 where MAC='"& mac &"'",1,3)
if  not rs.eof then
 if rs("id")<>id then
response.Write"<script> alert('该物理地址已经在系统购入记录中，请重新填写！'); history.back(); </script>"
response.end()
end if
end if
call closeRS(rs)
call createRs(rs,"select * from 电脑购入记录 where id='"& id &"'",1,3)
if not rs.eof then
mac_old=rs("MAC")
if buy_time<>"" then
rs("购入日期")=buy_time
end if
if buy_mon<>"" then
rs("购入单价")=Cdbl(buy_mon)
end if
rs("供应商")=buy_sup
rs("类型")=buy_mold
rs("位置")=place
rs("MAC")=mac
if zt="1" then
rs("状态")="报废"
end if
rs.update
end if
call closeRs(rs)
call createRS(rs,"select *from 电脑使用明细 where MAC='"& mac_old &"'",1,3)
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
call createRs(rsm,"select * from 电脑购入记录 where MAC='"& mac &"'",1,3)
if not rsm.eof then
    if rsm("状态")="闲置" or rsm("id")=buy_id then
        call createRs(rsa,"select * from 电脑使用明细 where id='"& use_id &"'",1,3)
        if not rsa.eof then
            rsa("使用部门")=use_dep
            rsa("责任人")=use_boss
            rsa("MAC")=mac
            rsa("开始使用时间")=use_timeone
			if use_timetwo<>"" then
            rsa("结束使用时间")=use_timetwo
			end if
			rsa("ipdress")=ip
			rsa("双IP")=db_ip
			rsa("办公室")=place
			rsa("备注")=note
            rsa.update
        end if
        call closeRS(rsa)
        response.Redirect("computer03.asp")
    else
        response.Write"<script> alert('标记此物理地址的电脑正在使用，请重新选择！'); history.back(); </script>"
        response.end()
    end if
else
    response.Write"<script> alert('系统中无此电脑，请重新选择！'); history.back(); </script>"
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
call createRs(rs,"select * from 电脑维护记录 where id='"& id &"'",1,3)
if not rs.eof then
rs("MAC")=fc
rs("comuser")=fm
rs("维护人")=it
rs("维护项目")=thing
rs("维护原因")=reason
rs("备注")=bz
rs("维护日期")=Cdate(ft)
rs("现象描述")=xx
rs("IT意见")=im
rs("依赖时间")=sd
rs("状态")=zt
rs.update()
end if
call closeRS(rs)
response.Redirect("computer04.asp")
end sub

sub delzero()
id=Trim(request.QueryString("id"))
call createRS(rs,"select * from 电脑购入记录 where id='"& id &"'",1,3)
if not rs.eof then
rs.delete
rs.update
end if
call closeRs(rs)
response.Redirect("computer02.asp")
end sub

sub delone()
id=Trim(request.QueryString("id"))
call createRS(rs,"select * from 电脑使用明细 where id='"& id &"'",1,3)
if not rs.eof then
mac=rs("MAC")
rs("结束使用时间")=date()
rs.update
call createRs(rsm,"select * from 电脑购入记录 where MAC='"& mac &"'",1,3)
if not rsm.eof then
rsm("状态")="闲置"
rsm("位置")=rs("办公室")
rsm.update
end if
call closeRs(rsm)
end if
call closeRs(rs)
response.Redirect("computer03.asp")
end sub

sub deltwo()
id=Trim(request.QueryString("id"))
call createRS(rs,"select * from 电脑维护记录 where id='"& id &"'",1,3)
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
response.Write"<script> alert('系统中无此用户，请重新选择！'); history.back(); </script>"
response.end()
end if
call closeRs(rs)
response.Write"<script> alert('添加成功'); window.location.href = 'eamiladd.asp';</script>"
end sub

call closeConn()
%>






