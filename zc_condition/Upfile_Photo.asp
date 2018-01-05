<%@language=vbscript codepage=936 %>
<!--#include file="Inc/config.asp"-->
<!--#include file="Inc/upfile_class.asp"-->
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>文件上传成功</title>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
<%
const upload_type=0   '上传方法：0=无惧无组件上传类，1=FSO上传 2=lyfupload，3=aspupload，4=chinaaspupload

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
	response.write "系统未开放文件上传功能"
else
	select case upload_type
			case 0
				call upload_0()  '使用化境无组件上传类
			case else
				'response.write "本系统未开放插件功能"
				'response.end
		end select
end if
%>

<%
sub upload_0()    '使用化境无组件上传类
	set upload=new upfile_class ''建立上传对象
	upload.GetData(104857600)   '取得上传数据,限制最大上传100M
	if upload.err > 0 then  '如果出错
		select case upload.err
			case 1
				response.write "请先选择你要上传的文件！"
			case 2
				response.write "你上传的文件总大小超出了最大限制（100M）"
		end select
		response.end
	end if
		SavePath = SaveUpFilesPath   '存放上传文件的目录
	if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
		
	for each formName in upload.file '列出所有上传了的文件
		set ofile=upload.file(formName)  '生成一个文件对象
		oFileSize=ofile.filesize
		if oFileSize<100 then
			msg="请先选择你要上传的文件！"
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
			msg="这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
			FoundErr=true
		end if
		
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&Trim(upload.Form("QA_ID"))&"-"&Trim(upload.Form("QA_name"))&"-"&ranNum&"."&fileExt

			ofile.SaveToFile Server.mappath(FileName)   '保存文件			
				call CreateRs(rs,"select * from zs_condition",1,3)
				rs.addnew
				rs("zx_number")=Trim(upload.Form("QA_ID"))				
				rs("zx_jt")=UCase(Trim(upload.Form("QA_name")))
				rs("zx_Address")=filename				
				rs("uptime")=date()
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

