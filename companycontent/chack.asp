<!--#include file="../connet/conn.asp" -->
<% 
call CreateRs(rs,"select * from company where id="&Trim(Request.QueryString("id")),1,3)
rs("folg")=0
rs.update
call closeRs(rs)
call closeConn()
response.Redirect("addparent.asp")
 %>