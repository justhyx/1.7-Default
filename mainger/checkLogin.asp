<!--#include file="../connet/sconn.asp" -->
<%
if Trim(Request.QueryString("ee"))="modify" then
uid=Trim(Request.Form("uid"))
opwd=Trim(Request.Form("opwd"))
npwd=Trim(Request.Form("npwd"))
kpwd=Trim(Request.Form("kpwd"))
	if npwd <> kpwd then
		response.Redirect("modify_pwd.asp?txtmage=��������ȷ�����벻һ��")
	end if
call CreateRs(rsm,"SELECT * FROM HUDSON_User INNER JOIN  department ON HUDSON_User.Department = department.code_id INNER JOIN  group_txt ON HUDSON_User.user_l = group_txt.group_id INNER JOIN manager_txt ON HUDSON_User.over_l = manager_txt.manager_id where HUDSON_User.UserID='"&uid&"' and HUDSON_User.UserPassword='"&opwd&"'",1,3)
	if rsm.eof then
		response.Redirect("modify_pwd.asp?txtmage=�û��������������ϸ����")
	end if
	rsm("UserPassword")=npwd
	rsm.update
call closeRs(rsm)
response.Write"<script> alert('�����޸ĳɹ�'); location.href= '../login.asp ';</script>"
response.End()
end if 
dim UserID,UserPassword,verifycode,UserName,UserSex,Department,Position,Phone,Tel,Power,addTime
UserID=Trim(Request.Form("UserID"))
UserPassword=Trim(Request.Form("pwd"))
verifycode=Trim(Request.Form("verifycode"))
call CreateRs(rs,"SELECT * FROM HUDSON_User INNER JOIN  group_txt ON HUDSON_User.user_l = group_txt.group_id INNER JOIN  manager_txt ON HUDSON_User.over_l = manager_txt.manager_id where HUDSON_User.UserID='"&UserID&"' and HUDSON_User.UserPassword='"&UserPassword&"'",1,1)
if rs.eof then
	response.Redirect("../login.asp?txtmage=�û��������������ϸ����")
end if
if rs("user_l")="" then
	response.Redirect("../login.asp?txtmage=��δ��ͨȨ������ϵ���²�")
end if
Session.Timeout=1440        'Session�Ự�ı���ʱ��
session("Power")=rs("Power")
session("user_id")=rs("UserID")
session("name_id")=rs("id")
session("UserName")=rs("UserName")
session("Department")=rs("Department")
session("group")=rs("position")
session("user_l")=rs("user_l")
session("over_l")=rs("over_l")
'if session("Url")="" then
response.Redirect("main.asp")
'else
'response.Redirect(session("Url"))
'end if
call closeRs(rs)
call closeConn()
 %>









