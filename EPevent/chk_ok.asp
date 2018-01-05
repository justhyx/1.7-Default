<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<% 
call CreateRs(rs,"select chk from Epevent where id='"& Trim(Request.QueryString("id")) &"'",1,3)
if rs("chk")="Î´Éó" then
	rs("chk")="ÒÑÉó"
else
	rs("chk")="Î´Éó"
end if
rs.update
call CloseRs(rs)
response.Redirect("eventlook.asp")
 %>