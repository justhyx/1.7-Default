<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<% 
call CreateRs(rs,"select chk from Epevent where id='"& Trim(Request.QueryString("id")) &"'",1,3)
if rs("chk")="δ��" then
	rs("chk")="����"
else
	rs("chk")="δ��"
end if
rs.update
call CloseRs(rs)
response.Redirect("eventlook.asp")
 %>