<!--#include file="UpFile_Class.asp" -->
<%
dim upfile,ServerPath,FSPath,oFile
upfilecount=0
set upfile=new upfile_class   '�����ϴ�����
upfile.AllowExt="Jpg"   '�����ϴ����͵ĺ�����
upfile.GetData (2048000) '�ϴ��ļ�������С
FSPath=GetFilePath(Server.mappath("./"),"\") 'ȡ�õ�ǰ�ļ��ڷ�����·�� 
Response.Write(FSPath&"<BR>")
ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/") 'ȡ������վ�ϵ�λ��
Response.Write(ServerPath&"<BR>")
for each formName in upfile.file '�г������ϴ��˵��ļ�
set oFile=upfile.file(formname)
Response.Write(TypeName(upfile.file)&"<BR>")
Response.Write upfile.AutoSave(formname,FSPath&oFile.FileName)   '�����ļ�
if upfile.iserr then 
   Response.Write upfile.errmessage
   Else
   Response.Write "�ϴ��ɹ�" 
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
