<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<% if session("UserName")="����" or session("UserName")="����"  then   %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("��ѡ��Ҫ�ϴ����ļ�");
		document.form1.FileName.focus();
    	return false;
  	}
}
</script>
<script language="javascript">
window.onload=function showtable(){
   var tablename=document.getElementById("mytable");
   var li=tablename.getElementsByTagName("tr");
   for (var i=0;i<=li.length;i++){
    if(i%2==0){
     li[i].style.backgroundColor="#F1F1F1";
     li[i].onmouseout=function(){
         this.style.backgroundColor="#F1F1F1";
      }
    }else{
     li[i].style.backgroundColor="#FFFFFF";
     li[i].onmouseout=function(){
         this.style.backgroundColor="#FFFFFF"
      }
    }
      li[i].onmouseover=function(){
      this.style.backgroundColor="#00CCFF";
      }
      
   }
}
</script>
</head>
<body>
<form id="form1" name="form1" method="post" action="Up.asp" onsubmit="return check()" enctype="multipart/form-data">
  <table width="100%" border="0" cellspacing="1" cellpadding="6">
    <tr>
      <td colspan="2" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="7%">������Ŀ</td>
      <td width="93%"><label>
        <select name="columnid" id="columnid">
          <option value="1">���ֲ�</option>
          <option value="2">����׼</option>
          <option value="3">���涨</option>
          <option value="4">&nbsp;&nbsp;��--��Ʒ������</option>
          <option value="5">&nbsp;&nbsp;��--Ӫҵ�ɹ���</option>
          <option value="6">&nbsp;&nbsp;��--�ֿ�����</option>
          <option value="7">&nbsp;&nbsp;��--���������</option>
          <option value="8">&nbsp;&nbsp;��--ע�ܼ���</option>
          <option value="9">&nbsp;&nbsp;��--���������</option>
          <option value="a">&nbsp;&nbsp;��--���ο�</option>
          <option value="b">&nbsp;&nbsp;��--������</option>
          <option value="c">&nbsp;&nbsp;��--�����</option>
		  <option value="e">&nbsp;&nbsp;��--IT������</option>
          <option value="d">&nbsp;&nbsp;��--������</option>
        </select>
        </label></td>
    </tr>
    <tr>
      <td>�ļ�����</td>
      <td><label>
        <input name="columncontent" type="text" id="columncontent" />
        </label></td>
    </tr>
    <tr>
      <td>�ļ��ϴ�</td>
      <td><input name="columnurl" type="file" id="columnurl" size="30" height="30" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="�� ��" />
        </label></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<table id="mytable" width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr>
    <td colspan="3" align="center" style="font-weight:bold">HUDSON�ļ�Ŀ¼</td>
  </tr>
  <tr style="font-weight:bold;">
    <td>Ŀ¼</td>
    <td>�ļ���</td>
    <td>����</td>
  </tr>
  <%  
'����ÿҳ��ʾ��������¼�ķ�ҳ����ThisPageSize
Const ThisPageSize=16
'�����ܼ�¼������ҳ������ǰҳ��������������
Dim ThisRsCount,ThisPageCount,ThisCurrentPage,i
'��URL�л�ȡҳ��Page�����ж���Ч��
ThisCurrentPage=Request("Page")
If Request("Page")="" Then
	ThisCurrentPage=1
ElseIf IsNumeric(ThisCurrentPage) Then
	ThisCurrentPage=int(ThisCurrentPage)
Else
	ThisCurrentPage=1
End If
If ThisCurrentPage<1 Then ThisCurrentPage=1
%>
  <%
'�����ѯ�ؼ��ֱ�������URL����л�ȡ��ѯ�Ĺؼ���TxtKey������ϳ�ģ����ѯSQL���
Dim TxtKey
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from company where folg='"&int(1)&"' Order By id DESC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from company Where columncontent like '%"&TxtKey&"%' Order By ID DESC",1,1)
		if rs.eof or rs.bof then
			response.Write("û�д˼�¼")
			response.End()
		end if
End If	
%>
  <%
'��ȡ�ܼ�¼��
ThisRsCount=Rs.RecordCount
'����ÿҳ��ʾ��������¼
Rs.PageSize=ThisPageSize
'��ȡ��ҳ��
ThisPageCount=Rs.PageCount
'�жϵ�ǰҳ�����Ч��
If ThisCurrentPage>ThisPageCount Then ThisCurrentPage=ThisPageCount
'���õ�ǰҳ��
Rs.AbsolutePage=ThisCurrentPage
%>
  <%
'ѭ����ȡ����
i=0
Do While Not Rs.EOF and i<ThisPageSize
%>
  <tr>
    <td>&nbsp;</td>
    <td><%= rs("columncontent") %>&nbsp;</td>
    <td><a href="chack.asp?id=<%= rs("id")%>">ɾ��</a></td>
  </tr>
  <%
	i=i+1
	Rs.MoveNext
Loop	
%>
</table>
<table width="760" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>�ܹ�<%= ThisRsCount %>����¼,<%= ThisPageSize %>����¼/ҳ,�ܹ�<%= ThisPageCount %>ҳ,��ǰ��<%= ThisCurrentPage %>ҳ</td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      ��һҳ
      <% Else %>
      <a href="?Page=1&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      ��һҳ
      <% Else %>
      <a href="?Page=<%=ThisCurrentPage-1%>&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      ��һҳ
      <% Else %>
      <a href="?Page=<%=ThisCurrentPage+1%>&TxtKey=<%=TxtKey%>" class="blue">��һҳ</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      ���һҳ
      <% Else %>
      <a href="?Page=<%=ThisPageCount%>&TxtKey=<%=TxtKey%>" class="blue">���һҳ</a>
      <% End If %>
    </td>
  </tr>
</table>
<% 
	else
	response.Write"<script> alert('ʹ���û���׼ȷ���뷵��'); history.back(); </script>"
	response.End()
	end if
 %>
</body>
</html>
<% call closeConn() %>
