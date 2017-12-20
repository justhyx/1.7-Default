<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/sconn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="refresh" content="180" >
<title>HUDSON工业管理系统--380T-2机台状态</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
	<% 
	dim bf,jhs,start_time,zq,s_zq,zzy,bb,scsl,qs,fs,s_overtime,scxl,ss,state_time
	call CreateRs(rs,"select * from equip where equip_name='380T-2'",1,1)
	select case rs("jt_state")
		case 0
			im="ok"
		case 1
			im="stop2"
			bg="FF0000"
		case 5
			im="j_t"
			bg="FFFF00"
	end select
	state_time=rs("state_time")
	call closeRs(rs)
	call CreateRs(rs,"select * from 作业日报 where 机台='380T-2' and work_state='Z'",1,1)
	if not rs.eof then	
	bf=rs("加工部番")
	jhs=rs("计划数量")
	start_time=rs("起始时间")
	zq=rs("周期")
	s_zq=rs("实际周期")
	zzy=rs("作业者")
	bb=rs("班别")	
	call closeRs(rs)
	call CreateRs(rs,"select * from 不良计算 where 部番='"&bf&"'",1,1)
	scsl=rs("计数")
	blzs=int(rs("黑点"))+int(rs("注塑不良"))+int(rs("银纹"))+int(rs("黑纹"))+int(rs("缩水"))+int(rs("划伤"))+int(rs("毛刺"))+int(rs("气泡"))+int(rs("断裂"))+int(rs("顶白"))+int(rs("变形"))+int(rs("碰伤"))+int(rs("脱模拉伤"))+int(rs("机械手跌落"))+int(rs("寸法不良"))+int(rs("其他"))	
	call closeRs(rs)
	call CreateRs(rs,"select * from goods where goods_name='"&bf&"'",1,1)  
	fs=rs("lb9_id")
	qs=rs("qs")
	call closeRs(rs)
	s_overtime=datediff("s",start_time,time())
	scxl=(cdbl(scsl)*cdbl(zq))/(cdbl(qs)*cdbl(s_overtime))*100
	ss=int((cdbl(jhsl)*cdbl(zq))/cdbl(qs))
	if scsl+blzs=0 then
		bll=0
	else
		bll= Formatnumber(blzs/(scsl+blzs)*100) '有可能除零
	end if
	scl= Formatnumber(scxl)
	bg_cr="FFFFFF"
	bg_img=im
	work_txt="作业时间"
	right_time=int(datediff("s",start_time,time()))
	left_time=int((cdbl(jhs)-cdbl(scsl))*cdbl(zq)/cdbl(qs))	
	else	
	bf="- -"
	jhs="- -"
	start_time="- -"
	zq="- -"
	s_zq="- -"
	zzy="- -"
	bb="- -"
	scsl="- -"
	fs="- -"
	qs="- -"
	s_overtime="- -"
	blzs="- -"
	bll="- -"
	scl="- -"
	ss="- -"
	bg_cr=bg
	bg_img=im
	work_txt="停机时间"
	left_time=0
	right_time=datediff("s",state_time,now())
	end if
'	call CreateRs(rs,"select * from prd_dictate where update_state='N' and equip_id='DA0000000130103' and dictate_date='"&date()&"'",1,1)
'		if rs.eof or rs.bof then
'			response.Write("没有计划")
'			response.End()
'		else
'			'rs.movenext
'			call CreateRs(rsa,"select * from goods where goods_id='"&rs("goods_id")&"'",1,1)
'			if not rs.eof then
'			bf_two=rsa("goods_name")
'			cl=rsa("cz")
'			else
'			bf="计划完成"
'			cl="- -"
'			end if			
'			call closeRs(rsa)
'			prd_qty=rs("prd_qty")
'			dictate_date=rs("dictate_date")
'		call closeRs(rs)			
'		end if	
	%>
<script type="text/javascript">
 var time=<%= left_time %>;
 var jtime=<%= right_time %>; 
 function j_time()
 	{ 		
		mtime=Math.round(time/60)
		document.getElementById("show").innerHTML=" 倒计时<font color=red>"+time+"</font>秒 | "+mtime+"分";
		time--;
		mjtime=Math.round(jtime/60)
		document.getElementById("z_show").innerHTML=" 已用时<font color=red>"+jtime+"</font>秒  | "+mjtime+"分";
		jtime++;
		window.setTimeout('j_time()',1000);
	} 
</script>
</head>

<body onload="j_time();">
<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="80" rowspan="8" align="center" bgcolor="<%= bg_cr %>"><img src="../img/<%= bg_img %>.gif" width="46" height="180" /></td>
    <td>部品番号</td>
    <td width="120"><%= bf %>&nbsp;</td>
  </tr>
  <tr>
    <td>计划数量</td>
    <td><%= jhs %></td>
  </tr>
  <tr>
    <td>生产取数</td>
    <td><%= qs %></td>
  </tr>
  <tr>
    <td>生产数量</td>
    <td><%= scsl %></td>
  </tr>
  <tr>
    <td>不良数量</td>
    <td><%= blzs %></td>
  </tr>
  <tr>
    <td>不 良 率</td>
    <td><%= bll %> &nbsp;</td>
  </tr>
  <tr>
    <td>生 产 率</td>
    <td><%= scl %></td>
  </tr>
  <tr>
    <td>开始时间</td>
    <td><%= start_time %></td>
  </tr>
  <tr>
    <td align="center">380T-2</td>
    <td><%= work_txt %></td>
    <td align="center" id="z_show">&nbsp;</td>
  </tr>
  <tr>
    <td><%= bf_two %></td>
    <td>预离完成</td>
    <td id="show">&nbsp;</td>
  </tr>
  <tr>
    <td><%= prd_qty %></td>
    <td>理论周期</td>
    <td><%= zq %></td>
  </tr>
  <tr>
    <td><%= dictate_date %></td>
    <td>实际周期</td>
    <td><%= s_zq %></td>
  </tr>
  <tr>
    <td><%= cl %></td>
    <td>作 业 员</td>
    <td><%= zzy %></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>作业班别</td>
    <td><%= bb %></td>
  </tr>
</table>
</body>
</html>
<% 
call closeConn()
 %>

