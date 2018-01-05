<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<SCRIPT language=javascript>
<!--#include file="../connet/conn.asp" -->
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("请选择要上传的文件");
		document.form1.FileName.focus();
    	return false;
  	}
}
</SCRIPT>
</head>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-left: 0px;
}
-->
</style>
<%
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i)
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function
%>
<link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css">
<body>
<form action="Upfile_Photo.asp" method="post" name="form1" onSubmit="return check()" enctype="multipart/form-data">
<table width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
          <tr>
            <td colspan="3" align="center" bgcolor="#86C4C4" style="color:#FFFFFF; font-weight:bold;">图档上传</td>
          </tr>
          <tr>
            <td width="20%" align="right">部品番号：</td>
            <td colspan="2" align="left"><label>
              <input name="QA_ID" type="text" id="QA_ID" />
            </label></td>
          </tr>
          <tr>
            <td align="right">机台：</td>
            <td colspan="2" align="left"><label>
              <input name="QA_name" type="text" id="QA_name" />
            </label></td>
          </tr>
          <tr>
            <td align="right">图档上传：</td>
            <td colspan="2" align="left"><input name="FileName" height="30" type="file" size="30" /></td>
          </tr>
          <tr>
            <td align="right">&nbsp;</td>
            <td colspan="2" align="left"><label></label></td>
          </tr>
          <tr>
            <td align="right">备注：</td>
            <td width="35%" align="left"><label>
              <textarea name="describe" rows="4" id="describe"></textarea>
            </label></td>
            <td width="45%" align="left" valign="bottom"><input type="submit" name="Submit" value="上 传" /></td>
          </tr>
  </table>

</form>
<table id="mytable" width="95%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
  <tr align="center">
    <td><strong>部品番号</strong></td>
    <td ><strong>机台</strong></td>
    <td ><strong>成形条件</strong></td>
    <td ><strong>登录时间</strong></td>
    <td ><strong>操作</strong></td>
  </tr>
  <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=10
'定义总记录数，总页数，当前页数，记数器变量
Dim ThisRsCount,ThisPageCount,ThisCurrentPage,i
'从URL中获取页码Page，并判断有效性
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
'定义查询关键字变量，从URL或表单中获取查询的关键字TxtKey，并组合成模糊查询SQL语句
Dim TxtKey,agree,ArrangeTime
TxtKey=Request("TxtKey")
If TxtKey="" Then
	call CreateRs(rs,"select * from zs_condition Order By id DESC",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from PicFiles Where name like '%"&TxtKey&"%' and okfolg='"&int(0)&"' Order by newfolg DESC",1,1)
			if rs.eof or rs.bof then
			response.Write("没有此部品")
			response.End()
		end if
End If	
%>
  <%
'获取总记录数
ThisRsCount=Rs.RecordCount
'设置每页显示多少条记录
Rs.PageSize=ThisPageSize
'获取总页数
ThisPageCount=Rs.PageCount
'判断当前页码的有效性
If ThisCurrentPage>ThisPageCount Then ThisCurrentPage=ThisPageCount
'设置当前页码
Rs.AbsolutePage=ThisCurrentPage
%>
  <%
'循环读取数据
i=0
Do While Not Rs.EOF and i<ThisPageSize
%>
  <tr align="center">
    <td><%= rs("zx_number") %>&nbsp;</td>
    <td><%= rs("zx_jt")%>&nbsp;</td>
    <td><a href="<%= rs("zx_address")%>"><%= rs("zx_number") %>成形条件</a>&nbsp;</td>
    <td><%= rs("uptime")%>&nbsp;</td>
    <td ><a href="chkok.asp?id=<%= rs("id")%>&amp;maseg=del" onclick="return confirm('您确定图档错误吗？点击确定将彻底删除本图档内容')">删除</a></td>
  </tr>
  <%
	i=i+1
	Rs.MoveNext	
Loop	
%>
</table>

<table width="760" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td>总共<%= ThisRsCount %>条记录,<%= ThisPageSize %>条记录/页,总共<%= ThisPageCount %>页,当前第<%= ThisCurrentPage %>页</td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      第一页
      <% Else %>
        <a href="?Page=1&amp;TxtKey=<%=TxtKey%>" class="blue">第一页</a>
        <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      上一页
      <% Else %>
        <a href="?Page=<%=ThisCurrentPage-1%>&amp;TxtKey=<%=TxtKey%>" class="blue">上一页</a>
        <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      下一页
      <% Else %>
        <a href="?Page=<%=ThisCurrentPage+1%>&amp;TxtKey=<%=TxtKey%>" class="blue">下一页</a>
        <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      最后一页
      <% Else %>
        <a href="?Page=<%=ThisPageCount%>&amp;TxtKey=<%=TxtKey%>" class="blue">最后一页</a>
        <% End If %>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</thml>