<%@language=vbscript codepage=936 %>
<!--#include file="Inc/config.asp"-->
<!--#include file="Inc/upfile_class.asp"-->
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ļ��ϴ��ɹ�</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<%
const upload_type=0   '�ϴ�������0=�޾�������ϴ��࣬1=FSO�ϴ� 2=lyfupload��3=aspupload��4=chinaaspupload

dim upload,oFile,formName,SavePath,filename,fileExt,oFileSize
dim EnableUpload
dim arrUpFileType
dim ranNum
dim msg,FoundErr
dim PhotoUrlID
msg=""
FoundErr=false
EnableUpload=false
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

</head>
<body leftmargin="2" topmargin="5" marginwidth="0" marginheight="0" >
<%
if EnableUploadFile="No" then
	response.write "ϵͳδ�����ļ��ϴ�����"
else
	select case upload_type
			case 0
				call upload_0()  'ʹ�û���������ϴ���
			case else
				'response.write "��ϵͳδ���Ų������"
				'response.end
		end select
end if
%>

<%
sub upload_0()    'ʹ�û���������ϴ���
	set upload=new upfile_class ''�����ϴ�����
	upload.GetData(104857600)   'ȡ���ϴ�����,��������ϴ�100M
	if upload.err > 0 then  '�������
		select case upload.err
			case 1
				response.write "����ѡ����Ҫ�ϴ����ļ���"
			case 2
				response.write "���ϴ����ļ��ܴ�С������������ƣ�100M��"
		end select
		response.end
	end if
		SavePath = SaveUpFilesPath   '����ϴ��ļ���Ŀ¼
	if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
		
	for each formName in upload.file '�г������ϴ��˵��ļ�
		set ofile=upload.file(formName)  '����һ���ļ�����
		oFileSize=ofile.filesize
		if oFileSize<100 then
			msg="����ѡ����Ҫ�ϴ����ļ���"
			FoundErr=True
		end if

		fileExt=lcase(ofile.FileExt)
		arrUpFileType=split(UpFileType,"|")
		for i=0 to ubound(arrUpFileType)
			if fileEXT=trim(arrUpFileType(i)) then
				EnableUpload=true
				exit for
			end if
		next
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" then
			EnableUpload=false
		end if
		if EnableUpload=false then
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
			FoundErr=true
		end if
		
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&Trim(upload.Form("QA_ID"))&"-"&Trim(upload.Form("QA_name"))&"-"&ranNum&"."&fileExt

			ofile.SaveToFile Server.mappath(FileName)   '�����ļ�			
				call CreateRs(rs,"select * from PicFiles",1,3)
				rs.addnew
				rs("Name")=UCase(Trim(upload.Form("QA_ID")))				
				rs("Suffix")=UCase(Trim(upload.Form("QA_name")))
				rs("DownAddress")=filename
				rs("uppeople")=session("UserName")
				rs("uptime")=now()
				rs("describe")=Trim(upload.Form("describe"))
				rs("newfolg")=1
				rs("okfolg")=1
				rs.update				
				call closeRs(rs)
				response.Redirect("Upload_Photo.asp")
		else

			strJS=strJS & "alert('" & msg & "');" & vbcrlf
		  	strJS=strJS & "history.go(-1);" & vbcrlf
		end if
		strJS=strJS & "</script>" & vbcrlf
		response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
call closeConn()

%>
</body>
</html>

