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
		response.Write"<script> alert('此时间段不能提单'); history.back()</script>"
		response.End()
	end if
elseif pd > 17 then
		response.Write"<script> alert('此时间段不能提单'); history.back()</script>"
		response.End()
end if
numberidadd=Trim(Request.Form("t2"))
call CreateRs(rs,"select * from trymolde",1,3)
rs.addnew
rs("品番")=numberidadd
rs("dress")=Request.ServerVariables("REMOTE_ADDR")
rs("addtime")=date()
rs.update
call closeRs(rs)
call CreateRs(rs,"select * from trymolde where 品番='"& numberidadd &"' and dress='"& Request.ServerVariables("REMOTE_ADDR") &"' and addtime='"& date() &"' order by id desc",1,3)
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
		response.Write"<script> alert('此时间段不能提单'); history.back()</script>"
		response.End()
	end if
elseif pdh > 17 then
		response.Write"<script> alert('此时间段不能提单'); history.back()</script>"
		response.End()
end if
id=Cint(request.form("tryid"))
radio=request.form("radiobutton")
call CreateRs(rs,"select * from trymolde where id='"& id &"'",1,3)
rs("试模类型")=Trim(Request.Form("s1"))
rs("生准要望机台")=Trim(Request.Form("t3"))
rs("穴数")=Trim(Request.Form("t5"))
rs("材料编号")=Trim(Request.Form("t6"))
rs("材料")=Trim(Request.Form("t7"))
rs("试作数量")=Trim(Request.Form("t8"))
rs("纳期")=Trim(Request.Form("t9"))
rs("模具状态")=Trim(Request.Form("t10"))
rs("备注")=Trim(Request.Form("b1"))
rs("模具担当")=Trim(Request.Form("t1"))
rs("用量")=Trim(Request.Form("t82"))
rs("用料F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("类型")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO处理")=request.Form("POchk")
rs("mold_type")=request.Form("mold_type")
rs("fzsb")=request.Form("fzsb")
rs.update
pf=rs("品番")
call closeRs(rs)
call createRs(rsa,"SELECT * FROM trymolde WHERE (模具担当 IS NULL) AND (试模类型 IS NULL) AND (纳期 IS NULL) ",1,3)
do while not rsa.eof
rsa.delete
rsa.update
rsa.movenext
loop
call closeRs(rsa)
if radio<>"" then
call createRs(rs,"select * from compose where 部品番号='"& pf &"'",1,3)
if rs.eof then 
rs.addnew
rs("部品番号")=pf
rs("试模类型")=Trim(Request.Form("s1"))
rs("机位")=Trim(Request.Form("t3"))
rs("穴数")=Trim(Request.Form("t5"))
rs("部品材料编号")=Trim(Request.Form("t6"))
rs("名称")=Trim(Request.Form("t7"))
rs("试作数量")=Trim(Request.Form("t8"))
rs("模具状态")=Trim(Request.Form("t10"))
rs("备注")=Trim(Request.Form("b1"))
rs("模具担当")=Trim(Request.Form("t1"))
rs("用量")=Trim(Request.Form("t82"))
rs("用料F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("类型")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO处理")=request.Form("POchk")
rs.update
else
rs("试模类型")=Trim(Request.Form("s1"))
rs("机位")=Trim(Request.Form("t3"))
rs("穴数")=Trim(Request.Form("t5"))
rs("部品材料编号")=Trim(Request.Form("t6"))
rs("名称")=Trim(Request.Form("t7"))
rs("试作数量")=Trim(Request.Form("t8"))
rs("模具状态")=Trim(Request.Form("t10"))
rs("备注")=Trim(Request.Form("b1"))
rs("模具担当")=Trim(Request.Form("t1"))
rs("用量")=Trim(Request.Form("t82"))
rs("用料F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("类型")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO处理")=request.Form("POchk")
rs.update
end if
call closeRs(rs)
end if

response.Redirect("addtrymolde.asp?event=add")
end sub

sub exchange()
id=Cint(request.form("tryid"))
call CreateRs(rs,"select * from trymolde where id='"& id &"'",1,3)
rs("试模类型")=Trim(Request.Form("s1"))
rs("生准要望机台")=Trim(Request.Form("t3"))
rs("穴数")=Trim(Request.Form("t5"))
rs("材料编号")=Trim(Request.Form("t6"))
rs("材料")=Trim(Request.Form("t7"))
rs("试作数量")=Trim(Request.Form("t8"))
rs("纳期")=Trim(Request.Form("t9"))
rs("模具状态")=Trim(Request.Form("t10"))
rs("备注")=Trim(Request.Form("b1"))
rs("模具担当")=Trim(Request.Form("t1"))
rs("用量")=Trim(Request.Form("t82"))
rs("用料F")=Trim(Request.Form("t822"))
rs("okfolg")=1
rs("类型")=request.Form("t11")
rs("PO")=request.Form("t12")
rs("PO处理")=request.Form("POchk")
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
rs("生管确认机台")=Trim(Request.Form("okji"))
rs("安排时间")=Trim(Request.Form("arrangetime"))
rs("生管备注")=Trim(Request.Form("sgnote"))
rs.update
response.Redirect("arrange.asp")
call closeRs(rs)
end sub 

sub arrange_a()
call CreateRs(rs,"select * from trymolde where id="&Trim(Request.Form("id")),1,3)
rs("生管确认机台")=Trim(Request.Form("okji"))
rs("安排时间")=Trim(Request.Form("arrangetime"))
rs("生管备注")=Trim(Request.Form("sgnote"))
rs.update
response.Redirect("arrange_a.asp")
call closeRs(rs)
end sub

sub ment()
call CreateRs(rs,"select * from trymolde where id="&Trim(Request.Form("id")),1,3)
rs("模具状态")=Trim(Request.Form("moldements"))
rs.update
call closeRs(rs)
response.Redirect("statements.asp")
end sub
call closeConn()
 %>



