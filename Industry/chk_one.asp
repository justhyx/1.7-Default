<!--#include file="../connet/sconn.asp" -->
<% 
chk=Trim(Request.QueryString("event"))
select case chk
	case "ͣ��"
		call zs_stop()
	case "����"
		call zs_xw()
	case "����"
		call zs_sc()
	case "�쳣"
		call zs_yc()
	case "start"
		call start()
	case "����ͳ��"
		call zs_sctj()
	case "xianwai"
		call xianwai()
	case "ͣ��ԭ��"
		call zs_stopwhy()
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
response.write"<script> alert('�����Ϣ��д������'); history.back();</script>"
response.end()
end if

call CreateRs(rs,"select * from ��������",1,3)       'ͳ����Ϣ��ʼ��
rs.addnew
rs("����")=Trim(Request.Form("bf"))
rs("��̨")=Trim(Request.Form("jt"))
rs("ʱ��")=date()
rs("�ڵ�")=0
rs("ע�ܲ���")=0
rs("����")=0
rs("����")=0
rs("��ˮ")=0
rs("����")=0
rs("ë��")=0
rs("����")=0
rs("����")=0
rs("����")=0
rs("����")=0
rs("����")=0
rs("��ģ����")=0
rs("��е�ֵ���")=0
rs("�編����")=0
rs("����")=0
rs("����")=0
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from ��ҵ�ձ�",1,3)
randomize
ranNuma=int(9000*rnd)+100
rs.addnew
rs("ID")="TJ"&year(date())&month(date())&day(date())&ranNuma
rs("��ʼʱ��")=time()
rs("��̨")=request.cookies("scjt")
rs("�ӹ�����")=request.cookies("bf")
rs("������")=request.cookies("dd")
rs("��ҵ��")=request.cookies("zyy")
rs("��ҵ����")=date()
rs("���")=request.cookies("bb")
rs("����")=Trim(Request.Form("zq"))
rs.update
call closeRs(rs)
response.Redirect("dbl_one_sc.asp")
end sub
 '*********ͣ������**********************  
sub zs_stop()
randomize
ranNum=int(9000*rnd)+100
call CreateRs(rs,"select * from ͣ��ʱ��",1,3)
rs.addnew
rs("ID")=TJ&year(date())&month(date())&day(date())&ranNum
rs("����")=request.cookies("bf")
rs("��̨")=request.cookies("scjt")
rs("��ҵ����")=date()
rs("���")=request.cookies("bb")
rs("������")=request.cookies("dd")
rs("��ʼʱ��")=time()
rs.update
call closeRs(rs)
if request.Cookies("bf_one")<>"" then
randomize
ranNumone=int(9000*rnd)+100
call CreateRs(rs,"select * from ͣ��ʱ��",1,3)
rs.addnew
rs("ID")=TJ&year(date())&month(date())&day(date())&ranNumone
rs("����")=request.cookies("bf_one")
rs("��̨")=request.cookies("scjt")
rs("��ҵ����")=date()
rs("���")=request.cookies("bb")
rs("������")=request.cookies("dd")
rs("��ʼʱ��")=time()
rs.update
call closeRs(rs)
end if
call CreateRs(rs,"select * from ��λ״̬ where ��λ='"&request.cookies("scjt")&"'",1,3)           '��̨״̬����
rs("״̬")="3"                                                                                    '��̨״̬����
rs.update
call closeRs(rs)
response.Redirect("zs_xw.asp")
end sub                                                               
'************���⵽��ʱ��*******************
sub xianwai()
call CreateRs(rs,"select * from ͣ��ʱ�� where ��̨='"&request.cookies("scjt")&"' and ��ҵ����='"&date()&"' and ����='"&request.cookies("bf")&"'",1,3)
rs("����ʱ��")=time()
rs.update
call closeRs(rs)
if request.Cookies("bf_one")<>"" then
call CreateRs(rs,"select * from ͣ��ʱ�� where ��̨='"&request.cookies("scjt")&"' and ��ҵ����='"&date()&"' and ����='"&request.cookies("bf_one")&"'",1,3)
rs("����ʱ��")=time()
rs.update
call closeRs(rs)
end if
response.Redirect("zs_stop.asp")
end sub
'************ͣ��ԭ��*******************
sub zs_stopwhy()
call CreateRs(rs,"select * from ͣ��ʱ�� where ��̨='"&request.cookies("scjt")&"' and ��ҵ����='"&date()&"' and ����='"&request.cookies("bf")&"'",1,3)
rs("ͣ��ԭ��")=Trim(Request.QueryString("xinxi"))
rs.update
call closeRs(rs)
if request.Cookies("bf_one")<>"" then
call CreateRs(rs,"select * from ͣ��ʱ�� where ��̨='"&request.cookies("scjt")&"' and ��ҵ����='"&date()&"' and ����='"&request.cookies("bf")&"'",1,3)
rs("ͣ��ԭ��")=Trim(Request.QueryString("xinxi"))
rs.update
end if
response.redirect("dbl_thr_sc.asp")
end sub
'************����ʱ���ռ��Լ���̨״̬����*******************
sub zs_kaiji()
call CreateRs(rs,"select * from ͣ��ʱ�� where ��̨='"&request.cookies("scjt")&"' and ��ҵ����='"&date()&"' and ����='"&request.cookies("bf")&"'",1,3)
rs("��ֹʱ��")=time()
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from ��λ״̬ where ��λ='"&request.cookies("scjt")&"'",1,3)             '��̨״̬����
rs("״̬")="0"                                                                   '��ҵԱͣ����״̬Ϊδ֪ͣ��״̬
rs.update
call closeRs(rs)
response.redirect("dbl_thr_sc.asp")
end sub
'ͳ�Ʋ�����Ʒ��ʱ��
sub zs_sc()
	bottn=Trim(Request.QueryString("bottn"))	
	call CreateRs(rs,"select * from �������� where ����='"&request.cookies("bf")&"' and ʱ��='"&date()&"'",1,3)
	select case bottn
	case "a"
		rs("�ڵ�")=rs("�ڵ�")+1
	case "b"
		rs("ע�ܲ���")=rs("ע�ܲ���")+1
	case "c"
		rs("����")=rs("����")+1
	case "d"
		rs("����")=rs("����")+1
	case "e"
		rs("��ˮ")=rs("��ˮ")+1
	case "f"
		rs("����")=rs("����")+1
	case "g"
		rs("ë��")=rs("ë��")+1
	case "h"
		rs("����")=rs("����")+1
	case "i"
		rs("����")=rs("����")+1
	case "j"
		rs("����")=rs("����")+1
	case "k"
		rs("����")=rs("����")+1
	case "l"
		rs("����")=rs("����")+1
	case "m"
		rs("��ģ����")=rs("��ģ����")+1
	case "n"
		rs("��е�ֵ���")=rs("��е�ֵ���")+1
	case "p"
		rs("�編����")=rs("�編����")+1
	case "q"
		rs("����")=rs("����")+1
	end select
rs.update
call closeRs(rs)
response.Redirect("dbl_one_sc.asp")
end sub
'********************************
'������������
'********************************
sub zs_js()
	call CreateRs(rs,"select * from �������� where ʱ��='"&date()&"' and ��̨='"&request.cookies("scjt")&"'",1,3)
	rs("����")=int(rs("����"))+int(request.cookies("zxs"))
	rs.update
	call closeRs(rs)
	response.Redirect("dbl_one_sc.asp")
end sub
 '*************��������ͳ��******
sub zs_sctj()                 
	randomize
	ranNumuu=int(9000*rnd)+100
	call CreateRs(rs,"select * from �������� where ����='"&request.cookies("bf")&"' and ��̨='"&request.cookies("scjt")&"' and ʱ��='"&date()&"'",1,1)	
	if rs("�ڵ�")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="a"
			rsj("����")=rs("�ڵ�")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("ע�ܲ���")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="b"
			rsj("����")=rs("ע�ܲ���")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="c"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="d"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("��ˮ")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="e"
			rsj("����")=rs("��ˮ")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="f"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("ë��")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="g"
			rsj("����")=rs("ë��")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="h"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="i"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="j"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="k"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="l"
			rsj("����")=rs("����")			
			rsj.update
			call closeRs(rsj)
	end if
	if rs("��ģ����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="m"
			rsj("����")=rs("��ģ����")			
			rsj.update
			call closeRs(rsj)
	end if			
	if rs("��е�ֵ���")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="n"
			rsj("����")=rs("��е�ֵ���")			
			rsj.update
			call closeRs(rsj)
	end if						
	if rs("�編����")>0 then
		call CreateRs(rsj,"select * from ��������ͳ��",1,3)
			rsj.addnew
			rsj("���")=year(date())&month(date())&day(date())&ranNumuu  
			rsj("��������")=request.cookies("bf")
			rsj("������̨")=request.cookies("scjt")
			rsj("��������")=date()
			rsj("���")=request.cookies("bb")
			rsj("��������")="p"
			rsj("����")=rs("�編����")			
			rsj.update
			call closeRs(rsj)
	end if
	call closeRs(rs)
	'��¼���ͣ��ʱ��
	call CreateRs(rs,"select * from ��ҵ�ձ� where ��̨='"&request.cookies("scjt")&"' and �ӹ�����='"&request.cookies("bf")&"' and ��ҵ����='"&date()&"'",1,3)
	rs("��ֹʱ��")=time()
	rs.update
	zs_ztime=datediff("n",cdate(rs("��ʼʱ��")),cdate(rs("��ֹʱ��")))   '��/����ʱ��
	if zs_ztime<0 then	
	zs_ztimea=Cdbl(24+FormatNumber(zs_ztime/60,2,-1,0,0))
	else	
	zs_ztimea=Cdbl(FormatNumber(zs_ztime/60,2,-1,0,0))	
	end if
	call closeRs(rs)
	 '��ʼ���ƻ��װ嵱ǰ��Ʒ���״̬                                     
	call CreateRs(rs,"select * from �ƻ��װ� where equip_name='"&request.cookies("scjt")&"' and goods_name='"&request.cookies("bf")&"'",1,3)
	rs("state")="Y"
	rs.update
	call closeRs(rs)	
	'�����̨ͣ��ʱ�����
	call CreateRs(rs,"select * from ͣ��ʱ�� where ��̨='"&request.cookies("scjt")&"' and ��ҵ����='"&date()&"' and ����='"&request.cookies("bf")&"'",1,3)
	if not rs.eof then
	if rs("��ʼʱ��")<>"" then
	z_time=datediff("n",Cdate(rs("��ʼʱ��")),Cdate(rs("��ֹʱ��")))
	if z_time<0 then	
	rs("�ϼ�ʱ��")=Cdbl(24+FormatNumber(z_time/60,2,-1,0,0))
	else	
	rs("�ϼ�ʱ��")=Cdbl(FormatNumber(z_time/60,2,-1,0,0))	
	end if
	rs.update
	end if
	end if	
	call closeRs(rs)
	
'*************��������д����ҵ�ձ�****************
	'��������&����ͳ��
	call CreateRs(rs,"select * from �������� where ����='"&request.cookies("bf")&"' and ʱ��='"&date()&"' and ��̨='"&request.cookies("scjt")&"'",1,1) 
	blzs=int(rs("�ڵ�"))+int(rs("ע�ܲ���"))+int(rs("����"))+int(rs("����"))+int(rs("��ˮ"))+int(rs("����"))+int(rs("ë��"))+int(rs("����"))+int(rs("����"))+int(rs("����"))+int(rs("����"))+int(rs("����"))+int(rs("��ģ����"))+int(rs("��е�ֵ���"))+int(rs("�編����"))+int(rs("����"))
	zs_cl=int(rs("����"))
	call closeRs(rs)
	'���ʱ������
	call CreateRs(rs,"select * from ͣ��ʱ�� where ����='"&request.cookies("bf")&"' and ��ҵ����='"&date()&"' and ��̨='"&request.cookies("scjt")&"'",1,1)
	if not rs.eof then
	dim zdhj
	zdhj=0
	do while not rs.eof
	zdhj=cdbl(zdhj)+cdbl(rs("�ϼ�ʱ��"))
	rs.movenext
	loop
	else
	zdhj=0
	end if
	call CreateRs(rs,"select * from ��ҵ�ձ� where ��̨='"&request.cookies("scjt")&"' and �ӹ�����='"&request.cookies("bf")&"' and ��ҵ����='"&date()&"'",1,3)
	rs("����")=int(zs_cl)
	rs("�ƻ�����")=request.cookies("jhsl") 	
	rs("������")=int(blzs)
	rs("�ϼ�ʱ��")=cdbl(zs_ztimea)
	rs("�ж�ʱ��")=cdbl(zdhj)
	rs.update
	rs("��׼ʱ��")=Formatnumber((cdbl(rs("����"))+cdbl(rs("������")))*rs("����")/(3600*cdbl((request.cookies("qs")))))
	rs.update
	rs("����Ч��")=Formatnumber(rs("��׼ʱ��")/(rs("�ϼ�ʱ��")-rs("�ж�ʱ��"))*100)
	rs.update	
	call closeRs(rs)
	'���������ʼ��
	call CreateRs(rs,"select * from �������� where ����='"&request.cookies("bf")&"' and ʱ��='"&date()&"' and ��̨='"&request.cookies("scjt")&"'",1,3)  
	do while not rs.eof
	rs.delete
	rs.update
	rs.movenext
	loop
	call closeRs(rs)
	response.Redirect("flishlist.asp")
end sub
	sub wxs()
	call CreateRs(rs,"select * from �������� where ����='"&request.cookies("bf")&"' and ʱ��='"&date()&"'",1,3)
	rs("����")=int(rs("����"))+int(Trim(Request.Form("wxs")))
	rs.update
	call closeRs(rs)
	response.Redirect("dbl_one_sc.asp")
	end sub
	sub cxzq_s()
	call CreateRs(rs,"select * from ��ҵ�ձ� where ��̨='"&request.cookies("scjt")&"' and �ӹ�����='"&request.cookies("bf")&"' and ��ҵ����='"&date()&"'",1,3)
	rs("ʵ������")=cdbl(Trim(Request.Form("cxzq_s")))
	rs.update
	call closeRs(rs)	
	response.Redirect("dbl_one_sc.asp")
	end sub
call closeConn()
 %>
