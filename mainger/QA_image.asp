<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<%
if session("UserName")="����" or session("UserName")="�ź�ƻ" then
if request.Form("action")="save" then
	call CreateRs(rs,"select * from QA_Check",1,3)
	rs.addnew
	rs("QA_ID")=Trim(Request.Form("QA_ID"))
	rs("QA_people")=session("UserName")
	rs("QA_addTime")=now()
	rs.update
	call closeRs(rs)
	response.Redirect("../up/Upload_Photo.asp")
	call closeConn()	
end if
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <p>&nbsp;</p>
  <table width="400" border="0" align="center" cellpadding="4" cellspacing="1">
    <tr>
      <td colspan="2" align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">QA/QC��Ʒ����ͼƬ�ϴ�</td>
    </tr>
    <tr>
      <td width="20%">��Ʒ��ţ�</td>
      <td width="80%"><label>
        <input name="QA_ID" type="text" id="QA_ID" size="25" />
      </label></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><label></label></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="ȷ ��" />
        <input name="action" type="hidden" id="action" value="save" />
      </label></td>
    </tr>
  </table>  
</form>
<% 
	else
	response.Write"<script> alert('ʹ���û���׼ȷ���뷵��'); history.back(); </script>"
	response.End()
end if
 %>
</body>
</html>
<script language="javascript">
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//����style=standard650,ֵ��������ʵ����Ҫ�޸�Ϊ������ʽ��,ͨ������ʽ�ĺ�̨�������ﵽ���������ϴ��ļ����ͼ��ļ���С
	var arr = showModalDialog("../Editor/dialog/i_upload.htm?style=standard650&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>