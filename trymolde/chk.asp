<!--#include file="../connet/conn.asp" -->
<%
action=Trim(Request.QueryString("action"))
select case action
case "addnumber"
call addnumber()
case "addcentent"
call addcent()
case "exchange"
call exchange()
case "modify"
call modify()
case "arrange"
call arrange()
case "arrange_a"
call arrange_a()
case "ments"
call ment()
end select

sub addnumber()
pdh=hour(now())
pdm=Minute(now()) 
if pd > 16 then
	if pdm >30 then
		response.Write"<script> alert('��ʱ��β����ᵥ'); history.back()</script>"
		response.End()
	end if
elseif pd > 17 then
		response.Write"<script> alert('��ʱ��β����ᵥ'); history.back()</script>"
		response.End()
end if
numberidadd=Trim(Request.Form("t2"))
call CreateRs(rs,"select * from trymolde",1,3)
rs.addnew
rs("Ʒ��")=numberidadd
rs("dress")=Request.ServerVariables("REMOTE_ADDR")
rs("addtime")=date()
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from trymolde where Ʒ��='"& numberidadd &"' and dress='"& Request.ServerVariables("REMOTE_ADDR") &"' and addtime='"& date() &"' order by id desc",1,3)
if not rs.eof then
nid=rs("id")
end if
call closeRs(rs)
response.Redirect("addtrymolde.asp?event=add&numberid="&nid&"")
end sub

sub addcent()
pdh=hour(now())
pdm=Minute(now()) 
if pdh > 16 then
	if pdm >30 then
		response.Write"<script> alert('��ʱ��β����ᵥ'); history.back()</script>"
		response.End()
	end if
elseif pdh > 17 then
		response.Write"<script> alert('��ʱ��β����ᵥ'); history.back()</script>"
		response.End()
end if
id=Cint(request.form("tryid"))
radio=request.form("radiobutton")
call CreateRs(rs,"select * from trymolde where id='"& id &"'",1,3)
rs("��ģ����")=Trim(Request.Form("s1"))
rs("��׼Ҫ����̨")=Trim(Request.Form("t3"))
rs("Ѩ��")=Trim(Request.Form("t5"))
rs("���ϱ��")=Trim(Request.Form("t6"))
rs("����")=Trim(Request.Form("t7"))
rs("��������")=Trim(Request.Form("t8"))
rs("����")=Trim(Request.Form("t9"))
rs("ģ��״̬")=Trim(Request.Form("t10"))
rs("��ע")=Trim(Request.Form("b1"))
rs("ģ�ߵ���")=Trim(Request.Form("t1"))
rs("����")=Trim(Request.Form("t82"))
rs("����F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("����")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO����")=request.Form("POchk")
rs("mold_type")=request.Form("mold_type")
rs("fzsb")=request.Form("fzsb")
rs.update
pf=rs("Ʒ��")
call closeRs(rs)
call createRs(rsa,"SELECT * FROM trymolde WHERE (ģ�ߵ��� IS NULL) AND (��ģ���� IS NULL) AND (���� IS NULL) ",1,3)
do while not rsa.eof
rsa.delete
rsa.update
rsa.movenext
loop
call closeRs(rsa)
if radio<>"" then
call createRs(rs,"select * from compose where ��Ʒ����='"& pf &"'",1,3)
if rs.eof then 
rs.addnew
rs("��Ʒ����")=pf
rs("��ģ����")=Trim(Request.Form("s1"))
rs("��λ")=Trim(Request.Form("t3"))
rs("Ѩ��")=Trim(Request.Form("t5"))
rs("��Ʒ���ϱ��")=Trim(Request.Form("t6"))
rs("����")=Trim(Request.Form("t7"))
rs("��������")=Trim(Request.Form("t8"))
rs("ģ��״̬")=Trim(Request.Form("t10"))
rs("��ע")=Trim(Request.Form("b1"))
rs("ģ�ߵ���")=Trim(Request.Form("t1"))
rs("����")=Trim(Request.Form("t82"))
rs("����F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("����")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO����")=request.Form("POchk")
rs.update
else
rs("��ģ����")=Trim(Request.Form("s1"))
rs("��λ")=Trim(Request.Form("t3"))
rs("Ѩ��")=Trim(Request.Form("t5"))
rs("��Ʒ���ϱ��")=Trim(Request.Form("t6"))
rs("����")=Trim(Request.Form("t7"))
rs("��������")=Trim(Request.Form("t8"))
rs("ģ��״̬")=Trim(Request.Form("t10"))
rs("��ע")=Trim(Request.Form("b1"))
rs("ģ�ߵ���")=Trim(Request.Form("t1"))
rs("����")=Trim(Request.Form("t82"))
rs("����F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("����")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO����")=request.Form("POchk")
rs.update
end if
call closeRs(rs)
end if

response.Redirect("addtrymolde.asp?event=add")
end sub

sub exchange()
id=Cint(request.form("tryid"))
call CreateRs(rs,"select * from trymolde where id='"& id &"'",1,3)
rs("��ģ����")=Trim(Request.Form("s1"))
rs("��׼Ҫ����̨")=Trim(Request.Form("t3"))
rs("Ѩ��")=Trim(Request.Form("t5"))
rs("���ϱ��")=Trim(Request.Form("t6"))
rs("����")=Trim(Request.Form("t7"))
rs("��������")=Trim(Request.Form("t8"))
rs("����")=Trim(Request.Form("t9"))
rs("ģ��״̬")=Trim(Request.Form("t10"))
rs("��ע")=Trim(Request.Form("b1"))
rs("ģ�ߵ���")=Trim(Request.Form("t1"))
rs("����")=Trim(Request.Form("t82"))
rs("����F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("����")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO����")=request.Form("POchk")
rs.update
call closeRs(rs)
response.Redirect("addtrymolde.asp?event=add")
end sub

sub modify() 
ipdress=Trim(Request.QueryString("ipdress"))
call CreateRs(rs,"select * from trymolde where id="&Trim(Request.QueryString("id")),1,3)
	if Trim(Request.QueryString("op"))="isok" then
	rs("okfolg")=0
		elseif Trim(Request.QueryString("op"))="isno" then
	rs.delete
	end if
	rs.update
call closeRs(rs)
response.Redirect(ipdress&".asp?event=add")
end sub

sub arrange()
call CreateRs(rs,"select * from trymolde where id="&Trim(Request.Form("id")),1,3)
rs("����ȷ�ϻ�̨")=Trim(Request.Form("okji"))
rs("����ʱ��")=Trim(Request.Form("arrangetime"))
rs("���ܱ�ע")=Trim(Request.Form("sgnote"))
rs.update
response.Redirect("arrange.asp")
call closeRs(rs)
end sub 

sub arrange_a()
call CreateRs(rs,"select * from trymolde where id="&Trim(Request.Form("id")),1,3)
rs("����ȷ�ϻ�̨")=Trim(Request.Form("okji"))
rs("����ʱ��")=Trim(Request.Form("arrangetime"))
rs("���ܱ�ע")=Trim(Request.Form("sgnote"))
rs.update
response.Redirect("arrange_a.asp")
call closeRs(rs)
end sub

sub ment()
call CreateRs(rs,"select * from trymolde where id="&Trim(Request.Form("id")),1,3)
rs("ģ��״̬")=Trim(Request.Form("moldements"))
rs.update
call closeRs(rs)
response.Redirect("statements.asp")
end sub
call closeConn()
 %>



