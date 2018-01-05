<!--#include file="../connet/sconn.asp" -->
<% 
chk=Trim(Request.QueryString("event"))
select case chk
	case "停机"
		call zs_stop()
	case "线外"
		call zs_xw()
	case "生产"
		call zs_sc()
	case "异常"
		call zs_yc()
	case "start"
		call start()
	case "生产统计"
		call zs_sctj()
	case "xianwai"
		call xianwai()
	case "停机原因"
		call zs_stopwhy()
	case "kaiji"
		call zs_kaiji()
	case "zc_js"
		call zs_js()
	case "wxs"
		call wxs()
	case "other_stop"
		call other_stop()
	case "cxqz_s"
		call cxzq_s()
	case "dq_in"
		call dq_in()
	case "tj_in"
		call tj_in()
	case "unusual"
		call unusual()
	end select
sub start()
response.End()
response.cookies("bf")=Trim(Request.Form("bf"))
response.Cookies("bf").Expires=dateadd("h",24,now())
response.cookies("zxs")=Trim(Request.Form("zxs"))
response.Cookies("zxs").Expires=dateadd("h",24,now())
response.cookies("cxzq")=Trim(Request.Form("zq"))
response.Cookies("cxzq").Expires=dateadd("h",24,now())
response.cookies("jhsl")=Trim(Request.Form("js"))
response.Cookies("jhsl").Expires=dateadd("h",24,now())
response.cookies("qs")=Trim(Request.Form("qs"))
response.Cookies("qs").Expires=dateadd("h",24,now())
response.Cookies("scjt")=Trim(Request.Form("jt"))
response.Cookies("scjt").Expires=dateadd("h",24,now())
response.Cookies("bb")=Trim(Request.Form("bb"))
response.Cookies("bb").Expires=dateadd("h",24,now())
response.Cookies("dd")=Trim(Request.Form("dd"))
response.Cookies("dd").Expires=dateadd("h",24,now())
response.Cookies("zyy")=Trim(Request.Form("zyy"))
response.Cookies("zyy").Expires=dateadd("h",24,now())
if request.cookies("bf")="" or request.cookies("scjt")="" or request.cookies("bb")="" or request.cookies("zxs")="" or Trim(Request.Form("zyy"))="" or request.cookies("dd")="" or Trim(Request.Form("js"))="" then
response.write"<script> alert('请把信息填写完整！'); history.back();</script>"
response.end()
end if
'机台状态改写
call CreateRs(rs,"select * from equip where equip_name='"&request.cookies("scjt")&"'",1,3)
rs("jt_state")="0"
rs("state_time")=now()
rs.update
call closeRs(rs)
'统计信息初始化
call CreateRs(rs,"select * from 不良计算",1,3)       
rs.addnew
rs("部番")=Trim(Request.Form("bf"))
rs("机台")=Trim(Request.Form("jt"))
rs("时间")=date()
rs("黑点")=0
rs("注塑不良")=0
rs("银纹")=0
rs("黑纹")=0
rs("缩水")=0
rs("划伤")=0
rs("毛刺")=0
rs("气泡")=0
rs("断裂")=0
rs("顶白")=0
rs("变形")=0
rs("碰伤")=0
rs("脱模拉伤")=0
rs("机械手跌落")=0
rs("寸法不良")=0
rs("其他")=0
rs("计数")=0
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from 作业日报",1,3)
randomize
ranNuma=int(9000*rnd)+100
rs.addnew
rs("ID")="TJ"&year(date())&month(date())&day(date())&ranNuma
rs("起始时间")=time()
rs("机台")=request.cookies("scjt")
rs("加工部番")=request.cookies("bf")
rs("管理担当")=request.cookies("dd")
rs("作业者")=request.cookies("zyy")
rs("作业日期")=date()
rs("班别")=request.cookies("bb")
rs("周期")=Trim(Request.Form("zq"))
rs("计划数量")=request.cookies("jhsl")
rs("work_state")="Z"
rs("prd_dictate_id")=request.cookies("prd_dictate_id")
rs.update
response.Cookies("zy_date")=rs("作业日期")
call closeRs(rs)
'***************开机时间*************
tt="真"
call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 备注='"&tt&"'",1,3)
if not rs.eof then
rs("终止时间")=time()
rs("备注")=""
rs.update
end if
call closeRs(rs)
response.Redirect("dq_in.asp")
end sub
 '*********停机处理**********************  
sub zs_stop()
randomize
ranNum=int(9000*rnd)+100
call CreateRs(rs,"select * from 停机时间",1,3)
rs.addnew
	rs("ID")=TJ&year(date())&month(date())&day(date())&ranNum
	rs("部番")=request.cookies("bf")
	rs("机台")=request.cookies("scjt")
	rs("作业日期")=date()
	rs("班别")=request.cookies("bb")
	rs("管理担当")=request.cookies("dd")
	rs("起始时间")=time()
	rs("备注")="wk"
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from equip where equip_name='"&request.cookies("scjt")&"'",1,3)           '机台状态处理
rs("jt_state")="1"                                                                                    '机台状态更新
rs.update
call closeRs(rs)
response.Redirect("zs_xw.asp")
end sub                                                               
'************线外到达时间*******************
sub xianwai()
if request.Cookies("bf")="" then
	call CreateRs(rs,"select * from 停机时间 where 备注='"&Request.ServerVariables("REMOTE_ADDR")&"'",1,3)
	rs("线外时间")=time()
	rs.update
	call closeRs(rs)
else
	call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&date()&"' and 部番='"&request.cookies("bf")&"'",1,3)
	rs("线外时间")=time()
	rs.update
	call closeRs(rs)
end if
response.Redirect("zs_stop.asp")
end sub
'************停机原因*******************
sub zs_stopwhy()
response.End()
if request.cookies("bf")="" then
	call CreateRs(rs,"select * from 停机时间 where 备注='"&Request.ServerVariables("REMOTE_ADDR")&"'",1,3)
	rs("停机原因")=Trim(Request.QueryString("xinxi"))
	rs("备注")="finish"
	rs.update
	response.Redirect("index.asp")
else
	response.End()
	call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&date()&"' and 部番='"&request.cookies("bf")&"' and 备注='wk'",1,3)
	rs("停机原因")=Trim(Request.QueryString("xinxi"))
	rs("备注")="kw"
	rs.update
end if
call closeRs(rs)
response.redirect("zs_start.asp")
end sub
'************开机时间收集以及机台状态更新*******************
sub zs_kaiji()
if request.cookies("bf")="" then
	response.Redirect("index.asp")
call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&date()&"' and 部番='"&request.cookies("bf")&"'",1,3)
rs("终止时间")=time()
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from equip where equip_name='"&request.cookies("scjt")&"'",1,3)   '机台状态处理
rs("jt_state")="0"                                                                           '作业员停机，状态为未知停机状态
rs.update
call closeRs(rs)
response.redirect("zs_sc.asp")
end sub
'统计不良部品临时表
sub zs_sc()
	bottn=Trim(Request.QueryString("bottn"))	
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf")&"' and 时间='"&request.Cookies("zy_date")&"'",1,3)
	select case bottn
	case "a"
		rs("黑点")=rs("黑点")+1
	case "b"
		rs("注塑不良")=rs("注塑不良")+1
	case "c"
		rs("银纹")=rs("银纹")+1
	case "d"
		rs("黑纹")=rs("黑纹")+1
	case "e"
		rs("缩水")=rs("缩水")+1
	case "f"
		rs("划伤")=rs("划伤")+1
	case "g"
		rs("毛刺")=rs("毛刺")+1
	case "h"
		rs("气泡")=rs("气泡")+1
	case "i"
		rs("断裂")=rs("断裂")+1
	case "j"
		rs("顶白")=rs("顶白")+1
	case "k"
		rs("变形")=rs("变形")+1
	case "l"
		rs("碰伤")=rs("碰伤")+1
	case "m"
		rs("脱模拉伤")=rs("脱模拉伤")+1
	case "n"
		rs("机械手跌落")=rs("机械手跌落")+1
	case "p"
		rs("寸法不良")=rs("寸法不良")+1
	case "q"
		rs("其他")=rs("其他")+1
	end select
rs.update
call closeRs(rs)
response.Redirect("zs_sc.asp")
end sub
'********************************
'生产数量计算
'********************************
sub zs_js()
	call CreateRs(rs,"select * from 不良计算 where 时间='"&request.Cookies("zy_date")&"' and 机台='"&request.cookies("scjt")&"'and 部番='"&request.Cookies("bf")&"'",1,3)
	rs("计数")=int(rs("计数"))+int(request.cookies("zxs"))
	rs.update
	call closeRs(rs)
	response.Redirect("zs_sc.asp")
end sub
 '*************不良内容统计******
sub zs_sctj()                 
	randomize
	ranNumuu=int(9000*rnd)+100
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf")&"' and 机台='"&request.cookies("scjt")&"' and 时间='"&request.Cookies("zy_date")&"'",1,1)	
	if rs("黑点")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="a"
			rsj("数量")=rs("黑点")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("注塑不良")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="b"
			rsj("数量")=rs("注塑不良")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("银纹")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="c"
			rsj("数量")=rs("银纹")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("黑纹")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="d"
			rsj("数量")=rs("黑纹")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("缩水")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="e"
			rsj("数量")=rs("缩水")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("划伤")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="f"
			rsj("数量")=rs("划伤")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("毛刺")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="g"
			rsj("数量")=rs("毛刺")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("气泡")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="h"
			rsj("数量")=rs("气泡")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("断裂")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="i"
			rsj("数量")=rs("断裂")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("顶白")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="j"
			rsj("数量")=rs("顶白")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("变形")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="k"
			rsj("数量")=rs("变形")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("碰伤")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="l"
			rsj("数量")=rs("碰伤")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("脱模拉伤")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="m"
			rsj("数量")=rs("脱模拉伤")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("机械手跌落")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="n"
			rsj("数量")=rs("机械手跌落")			
			rsj.update
			call closeRs(rsj)
	end if						
	if rs("寸法不良")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=request.Cookies("zy_date")
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="p"
			rsj("数量")=rs("寸法不良")			
			rsj.update
			call closeRs(rsj)
	end if
	call closeRs(rs)
	'记录完成停机时间
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"'",1,3)
	rs("终止时间")=time()
	rs.update
	zs_ztime=datediff("n",cdate(rs("起始时间")),cdate(rs("终止时间")))   '起/终总时间
	if zs_ztime<0 then	
	zs_ztimea=Cdbl(24+FormatNumber(zs_ztime/60,2,-1,0,0))
	else	
	zs_ztimea=Cdbl(FormatNumber(zs_ztime/60,2,-1,0,0))	
	end if
	call closeRs(rs)
	'计算机台停机时间计算
	call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&date()&"' and 部番='"&request.cookies("bf")&"'",1,3)
	if not rs.eof then
	if rs("起始时间")<>"" then
	z_time=datediff("n",Cdate(rs("起始时间")),Cdate(rs("终止时间")))
	if z_time<0 then	
	rs("合计时间")=Cdbl(24+FormatNumber(z_time/60,2,-1,0,0))
	else	
	rs("合计时间")=Cdbl(FormatNumber(z_time/60,2,-1,0,0))	
	end if
	rs.update
	end if
	end if	
	call closeRs(rs)	
'*************生产数量写入作业日报****************
	'不良总数&产量统计
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf")&"' and 时间='"&request.Cookies("zy_date")&"' and 机台='"&request.cookies("scjt")&"'",1,1) 
	blzs=int(rs("黑点"))+int(rs("注塑不良"))+int(rs("银纹"))+int(rs("黑纹"))+int(rs("缩水"))+int(rs("划伤"))+int(rs("毛刺"))+int(rs("气泡"))+int(rs("断裂"))+int(rs("顶白"))+int(rs("变形"))+int(rs("碰伤"))+int(rs("脱模拉伤"))+int(rs("机械手跌落"))+int(rs("寸法不良"))+int(rs("其他"))
	zs_cl=int(rs("计数"))
	call closeRs(rs)
	'间断时间总数
	call CreateRs(rs,"select * from 停机时间 where 部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"' and 机台='"&request.cookies("scjt")&"'",1,1)
	if not rs.eof then
	dim zdhj
	zdhj=0
	do while not rs.eof
	zdhj=cdbl(zdhj)+cdbl(rs("合计时间"))
	rs.movenext
	loop
	else
	zdhj=0
	end if
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"'",1,3)
	rs("产量")=int(zs_cl)
	rs("计划数量")=request.cookies("jhsl") 	
	rs("不良数")=int(blzs)
	rs("合计时间")=cdbl(zs_ztimea)
	rs("中断时间")=cdbl(zdhj)
	rs("work_state")="Y"
	prd_date_id=rs("prd_dictate_id")
	rs.update
	rs("标准时间")=Formatnumber((cdbl(rs("产量"))+cdbl(rs("不良数")))*rs("周期")/(3600*cdbl((request.cookies("qs")))))
	rs.update
	rs("生产效率")=Formatnumber(rs("标准时间")/(rs("合计时间")-rs("中断时间"))*100)
	rs.update	
	call closeRs(rs)
	'不良计算初始化
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf")&"' and 时间='"&request.Cookies("zy_date")&"' and 机台='"&request.cookies("scjt")&"'",1,3)  
	do while not rs.eof
	rs.delete
	rs.update
	rs.movenext
	loop
	call closeRs(rs)
	'机台状态改写
	call CreateRs(rs,"select * from equip where equip_name='"&request.cookies("scjt")&"'",1,3)
	rs("jt_state")="5"
	rs("state_time")=now()
	jt_id=rs("equip_id")
	rs.update
	call closeRs(rs)
	'记录停机时间
	randomize
	ranNum=int(9000*rnd)+100
	call CreateRs(rs,"select * from 停机时间",1,3)
	rs.addnew
	rs("ID")=TJ&year(date())&month(date())&day(date())&ranNum
	rs("部番")=request.cookies("bf")
	rs("机台")=request.cookies("scjt")
	rs("作业日期")=date()
	rs("班别")=request.cookies("bb")
	rs("管理担当")=request.cookies("dd")
	rs("起始时间")=time()
	rs("备注")=Request.ServerVariables("REMOTE_ADDR")
	rs.update
	call closeRs(rs)	
	'生产状态改写
	call CreateRs(rs,"select * from prd_dictate where dictate_id='"&request.Cookies("prd_dictate_id")&"' and equip_id='"&jt_id&"' and dictate_date='"&request.Cookies("zy_date")&"'",1,3)
	rs("update_state")="Y"
	rs.update
	call closeRs(rs)
	response.Redirect("flishlist.asp")
end sub
	sub wxs()
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf")&"' and 时间='"&request.Cookies("zy_date")&"'",1,3)
	rs("计数")=int(rs("计数"))+int(Trim(Request.Form("wxs")))
	rs.update
	call closeRs(rs)
	response.Redirect("zs_sc.asp")
	end sub
	sub other_stop()
	call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&request.Cookies("zy_date")&"' and 部番='"&request.cookies("bf")&"'",1,3)
	rs("停机原因")=Trim(Request.Form("xinxi"))
	rs.update
	call closeRs(rs)
	response.redirect("zs_start.asp")
	end sub
	sub cxzq_s()
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"'",1,3)
	rs("实际周期")=cdbl(Trim(Request.Form("cxzq_s")))
	rs.update
	call closeRs(rs)
	response.Redirect("zs_sc.asp")
	end sub
	sub dq_in()
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"'",1,3)
	rs("段取时间")=time()
	rs.update
	call closeRs(rs)
	response.redirect("tj_in.asp")
	end sub
	sub tj_in()
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"'",1,3)
	rs("调机时间")=time()
	rs.update
	call closeRs(rs)
	response.redirect("zs_sc.asp")	
	end sub
	sub unusual()
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&request.Cookies("zy_date")&"'",1,3)
	rs("实际周期")=0
	rs("产量")=0
	rs("不良数")=0
	rs("终止时间")=time()
	rs("合计时间")=0
	rs("中断时间")=0
	rs("生产效率")=0
	rs("标准时间")=0
	rs("备注")="异常"
	rs("work_state")="Y"
	rs.update
	call closeRs(rs)
	response.redirect("zs_over.asp")
	end sub
call closeConn()
 %>
