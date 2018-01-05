<!--#include file="../connet/conn.asp" -->
<% 
dim action
action=Trim(Request.QueryString("action"))
select case action
	case "adduser"
		call adduser()
	case "modify"
		call modify()
	case "del"
		call del()
end select
sub adduser()
if Trim(Request.Form("UserID"))="" or Trim(Request.Form("UserName"))="" then
response.Write"<script> alert('请信息填写完整'); history.back();</script>"
response.end()
elseif Trim(Request.Form("UserPassword"))<>Trim(Request.Form("UserPassword1")) then
response.Write"<script> alert('两次密码输入不一致'); history.back();</script>"
response.end()
end if
call CreateRs(rs,"select * from HUDSON_User",1,3)
rs.addnew
rs("UserID")=Trim(Request.Form("UserID"))
rs("UserPassword")=Trim(Request.Form("UserPassword"))
rs("UserName")=Trim(Request.Form("UserName"))
rs("UserSex")=Trim(Request.Form("UserSex"))
rs("Department")=Trim(Request.Form("Department"))
rs("Phone")=Trim(Request.Form("Phone"))
rs("Tel")=Trim(Request.Form("Tel"))
rs("Power")=Trim(Request.Form("Power"))
rs("addTime")=Trim(Request.Form("adddate"))
rs("user_l")=Trim(Request.Form("user_l"))
rs("over_l")=Trim(Request.Form("over_l"))
rs.update
call closeRs(rs)
response.redirect("userManger.asp")
end sub

sub modify()
if Trim(Request.Form("UserID"))="" or Trim(Request.Form("UserName"))="" then
response.Write"<script> alert('请信息填写完整'); history.back();</script>"
response.end()
elseif Trim(Request.Form("UserPassword"))<>Trim(Request.Form("UserPassword1")) then
response.Write"<script> alert('两次密码输入不一致'); history.back();</script>"
response.end()
end if
call CreateRs(rs,"select * from HUDSON_User where id="&Trim(Request.QueryString("id")),1,3)
rs("UserID")=Trim(Request.Form("UserID"))
rs("UserPassword")=Trim(Request.Form("UserPassword"))
rs("UserName")=Trim(Request.Form("UserName"))
rs("UserSex")=Trim(Request.Form("UserSex"))
rs("Department")=Trim(Request.Form("Department"))
rs("Phone")=Trim(Request.Form("Phone"))
rs("Tel")=Trim(Request.Form("Tel"))
rs("Power")=Trim(Request.Form("Power"))
rs("addTime")=Trim(Request.Form("adddate"))
rs("user_l")=Trim(Request.Form("user_l"))
rs("over_l")=Trim(Request.Form("over_l"))
rs.update
call closeRs(rs)
response.Write"<script> alert('修改成功'); window.location.href = 'UserManger.asp';</script>"
response.end()
end sub

sub del()
conn.execute("delete from HUDSON_User where id='"& Trim(Request.QueryString("id")) &"'")
response.Write"<script> alert('删除成功'); window.location.href = 'UserManger.asp';</script>"
response.end()
end sub

call closeConn()
 %>
