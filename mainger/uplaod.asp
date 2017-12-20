<!--#include file="UpFile_Class.asp" -->
<%
dim upfile,ServerPath,FSPath,oFile
upfilecount=0
set upfile=new upfile_class   '建立上传对象
upfile.AllowExt="Jpg"   '设置上传类型的黑名单
upfile.GetData (2048000) '上传文件的最大大小
FSPath=GetFilePath(Server.mappath("./"),"\") '取得当前文件在服务器路径 
Response.Write(FSPath&"<BR>")
ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/") '取得在网站上的位置
Response.Write(ServerPath&"<BR>")
for each formName in upfile.file '列出所有上传了的文件
set oFile=upfile.file(formname)
Response.Write(TypeName(upfile.file)&"<BR>")
Response.Write upfile.AutoSave(formname,FSPath&oFile.FileName)   '保存文件
if upfile.iserr then 
   Response.Write upfile.errmessage
   Else
   Response.Write "上传成功" 
    end if
set oFile=nothing
Next
set upfile=nothing
function GetFilePath(FullPath,str)
   If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
   Else
   GetFilePath = ""
   End If
End function
%>
