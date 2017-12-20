<!--#include file="../connet/conn.asp" -->
<% 
dim ID,PicName,Newfolg,maseg
maseg=Trim(Request.QueryString("maseg"))
ID=Trim(Request.QueryString("id"))
if maseg="del" then                             'É¾³ýÉÏ´«´íÎóÍ¼µµ
call CreateRs(rs,"select * from zs_condition where id="&ID,1,3)  
rs.delete
rs.update
call closeRs(rs)
response.Redirect("Upload_Photo.asp") 
end if
call closeConn()
 %>

