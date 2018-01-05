<!--#include file="../connet/conn.asp" -->
<% 
if Trim(Request.QueryString("objct"))="small" then
ID=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from QA_image where id="&ID,1,3)  
rs.delete
rs.update
call closeRs(rs)
response.Redirect("Upload_Photo.asp")
elseif Trim(Request.QueryString("objct"))="big" then
ID=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from QA_Check where id="&int(ID),1,3)
rs.delete
rs.update
call closeRs(rs)
elseif Trim(Request.QueryString("objct"))="big_list" then
ID=Trim(Request.QueryString("id"))
call CreateRs(rs,"select * from QA_Check where QA_ID='"&ID&"'",1,3)
rs.delete
rs.update
call closeRs(rs)
response.Redirect("../mainger/QA_image.asp")
end if
 %>






