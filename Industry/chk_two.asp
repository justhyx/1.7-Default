<!--#include file="../connet/sconn.asp" -->
<% 
chk=Trim(Request.QueryString("event"))
select case chk
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
	case "kaiji"
		call zs_kaiji()
	case "zc_js"
		call zs_js()
	case "wxs"
		call wxs()
	case "cxzq_s"
		call cxzq_s()
	end select
sub start()
response.cookies("bf_one")=Trim(Request.Form("bf"))
response.Cookies("bf_one").Expires=dateadd("h",24,now())
response.cookies("zxs_one")=Trim(Request.Form("zxs"))
response.Cookies("zxs_one").Expires=dateadd("h",24,now())
response.cookies("cxzq_one")=Trim(Request.Form("zq"))
response.Cookies("cxzq_one").Expires=dateadd("h",24,now())
response.cookies("jhsl_one")=Trim(Request.Form("js"))
response.Cookies("jhsl_one").Expires=dateadd("h",24,now())
response.cookies("qs_one")=Trim(Request.Form("qs"))
response.Cookies("qs_one").Expires=dateadd("h",24,now())
response.Cookies("scjt_one")=Trim(Request.Form("jt"))
response.Cookies("scjt_one").Expires=dateadd("h",24,now())
response.Cookies("bb_one")=Trim(Request.Form("bb"))
response.Cookies("bb_one").Expires=dateadd("h",24,now())
response.Cookies("dd")=Trim(Request.Form("dd"))
response.Cookies("dd").Expires=dateadd("h",24,now())
response.Cookies("zyy")=Trim(Request.Form("zyy"))
response.Cookies("zyy").Expires=dateadd("h",24,now())
if request.cookies("bf_one")="" or request.cookies("scjt_one")="" or request.cookies("bb_one")="" or request.cookies("zxs_one")="" or request.cookies("zyy")="" or request.cookies("dd")="" or Trim(Request.Form("js"))="" then
response.write"<script> alert('请把信息填写完整！'); history.back();</script>"
response.end()
end if

call CreateRs(rs,"select * from 不良计算",1,3)       '统计信息初始化
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
rs("加工部番")=request.cookies("bf_one")
rs("管理担当")=request.cookies("dd")
rs("作业者")=request.cookies("zyy")
rs("作业日期")=date()
rs("班别")=request.cookies("bb")
rs("周期")=Trim(Request.Form("zq"))
rs.update
call closeRs(rs)
response.Redirect("dbl_two_sc.asp")
end sub
'************开机时间收集以及机台状态更新*******************
sub zs_kaiji()
call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&date()&"' and 部番='"&request.cookies("bf_one")&"'",1,3)
rs("终止时间")=time()
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from 机位状态 where 机位='"&request.cookies("scjt")&"'",1,3)             '机台状态处理
rs("状态")="0"                                                                   '作业员停机，状态为未知停机状态
rs.update
call closeRs(rs)
response.redirect("dbl_two_sc.asp")
end sub
'统计不良部品临时表
sub zs_sc()
	bottn=Trim(Request.QueryString("bottn"))	
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf_one")&"' and 时间='"&date()&"'",1,3)
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
response.Redirect("dbl_two_sc.asp")
end sub
'********************************
'生产数量计算
'********************************
sub zs_js()
	call CreateRs(rs,"select * from 不良计算 where 时间='"&date()&"' and 机台='"&request.cookies("scjt")&"' and 部番='"&request.cookies("bf_one")&"'",1,3)
	rs("计数")=int(rs("计数"))+int(request.cookies("zxs"))
	rs.update
	call closeRs(rs)
	response.Redirect("dbl_two_sc.asp")
end sub
 '*************不良内容统计******
sub zs_sctj()                 
	randomize
	ranNumuu=int(9000*rnd)+100
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf_one")&"' and 机台='"&request.cookies("scjt")&"' and 时间='"&date()&"'",1,1)	
	if rs("黑点")>0 then
		call CreateRs(rsj,"select * from 不良内容统计",1,3)
			rsj.addnew
			rsj("编号")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
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
			rsj("生产部番")=request.cookies("bf_one")
			rsj("生产机台")=request.cookies("scjt")
			rsj("生产日期")=date()
			rsj("班别")=request.cookies("bb")
			rsj("不良代码")="p"
			rsj("数量")=rs("寸法不良")			
			rsj.update
			call closeRs(rsj)
	end if
	call closeRs(rs)
	'记录完成停机时间
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf_one")&"' and 作业日期='"&date()&"'",1,3)
	rs("终止时间")=time()
	rs.update
	zs_ztime=datediff("n",cdate(rs("起始时间")),cdate(rs("终止时间")))   '起/终总时间
	if zs_ztime<0 then	
	zs_ztimea=Cdbl(24+FormatNumber(zs_ztime/60,2,-1,0,0))
	else	
	zs_ztimea=Cdbl(FormatNumber(zs_ztime/60,2,-1,0,0))	
	end if
	call closeRs(rs)
	 '初始化计划白板当前部品完成状态                                     
	call CreateRs(rs,"select * from 计划白板 where equip_name='"&request.cookies("scjt")&"' and goods_name='"&request.cookies("bf_one")&"'",1,3)
	rs("state")="Y"
	rs.update
	call closeRs(rs)	
	'计算机台停机时间计算
	call CreateRs(rs,"select * from 停机时间 where 机台='"&request.cookies("scjt")&"' and 作业日期='"&date()&"' and 部番='"&request.cookies("bf_one")&"'",1,3)
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
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf_one")&"' and 时间='"&date()&"' and 机台='"&request.cookies("scjt")&"'",1,1) 
	blzs=int(rs("黑点"))+int(rs("注塑不良"))+int(rs("银纹"))+int(rs("黑纹"))+int(rs("缩水"))+int(rs("划伤"))+int(rs("毛刺"))+int(rs("气泡"))+int(rs("断裂"))+int(rs("顶白"))+int(rs("变形"))+int(rs("碰伤"))+int(rs("脱模拉伤"))+int(rs("机械手跌落"))+int(rs("寸法不良"))+int(rs("其他"))
	zs_cl=int(rs("计数"))
	call closeRs(rs)
	'间断时间总数
	call CreateRs(rs,"select * from 停机时间 where 部番='"&request.cookies("bf_one")&"' and 作业日期='"&date()&"' and 机台='"&request.cookies("scjt")&"'",1,1)
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
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf_one")&"' and 作业日期='"&date()&"'",1,3)
	rs("产量")=int(zs_cl)
	rs("计划数量")=request.cookies("jhsl") 	
	rs("不良数")=int(blzs)
	rs("合计时间")=cdbl(zs_ztimea)
	rs("中断时间")=cdbl(zdhj)
	rs.update
	rs("标准时间")=Formatnumber((cdbl(rs("产量"))+cdbl(rs("不良数")))*rs("周期")/(3600*cdbl((request.cookies("qs")))))
	rs.update
	rs("生产效率")=Formatnumber(rs("标准时间")/(rs("合计时间")-rs("中断时间"))*100)
	rs.update	
	call closeRs(rs)
	'不良计算初始化
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf_one")&"' and 时间='"&date()&"' and 机台='"&request.cookies("scjt")&"'",1,3)  
	do while not rs.eof
	rs.delete
	rs.update
	rs.movenext
	loop
	call closeRs(rs)
	response.Redirect("flishlist.asp")
end sub
	sub wxs()
	call CreateRs(rs,"select * from 不良计算 where 部番='"&request.cookies("bf_one")&"' and 时间='"&date()&"'",1,3)
	rs("计数")=int(rs("计数"))+int(Trim(Request.Form("wxs")))
	rs.update
	call closeRs(rs)
	response.Redirect("dbl_two_sc.asp")
	end sub
	sub cxzq_s()
	call CreateRs(rs,"select * from 作业日报 where 机台='"&request.cookies("scjt")&"' and 加工部番='"&request.cookies("bf")&"' and 作业日期='"&date()&"'",1,3)
	rs("实际周期")=cdbl(Trim(Request.Form("cxzq_s")))
	rs.update
	call closeRs(rs)	
	response.Redirect("dbl_two_sc.asp")
	end sub
call closeConn()
 %>
