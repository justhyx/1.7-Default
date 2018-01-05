<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<% if session("UserName")="吴美" or session("UserName")="王丽"  then   %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<script language="javascript">
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
      <td width="7%">所属栏目</td>
      <td width="93%"><label>
        <select name="columnid" id="columnid">
          <option value="1">├手册</option>
          <option value="2">├标准</option>
          <option value="3">└规定</option>
          <option value="4">&nbsp;&nbsp;├--部品技术课</option>
          <option value="5">&nbsp;&nbsp;├--营业采购课</option>
          <option value="6">&nbsp;&nbsp;├--仓库管理课</option>
          <option value="7">&nbsp;&nbsp;├--生产管理课</option>
          <option value="8">&nbsp;&nbsp;├--注塑检查课</option>
          <option value="9">&nbsp;&nbsp;├--人事总务课</option>
          <option value="a">&nbsp;&nbsp;├--成形课</option>
          <option value="b">&nbsp;&nbsp;├--生产课</option>
          <option value="c">&nbsp;&nbsp;├--财务课</option>
		  <option value="e">&nbsp;&nbsp;├--IT技术课</option>
          <option value="d">&nbsp;&nbsp;└--技术课</option>
        </select>
        </label></td>
    </tr>
    <tr>
      <td>文件名称</td>
      <td><label>
        <input name="columncontent" type="text" id="columncontent" />
        </label></td>
    </tr>
    <tr>
      <td>文件上传</td>
      <td><input name="columnurl" type="file" id="columnurl" size="30" height="30" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="上 传" />
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
    <td colspan="3" align="center" style="font-weight:bold">HUDSON文件目录</td>
  </tr>
  <tr style="font-weight:bold;">
    <td>目录</td>
    <td>文件名</td>
    <td>操作</td>
  </tr>
  <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=16
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
			response.Write("没有此记录")
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
  <tr>
    <td>&nbsp;</td>
    <td><%= rs("columncontent") %>&nbsp;</td>
    <td><a href="chack.asp?id=<%= rs("id")%>">删除</a></td>
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
      <a href="?Page=1&TxtKey=<%=TxtKey%>" class="blue">第一页</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=1 Then %>
      上一页
      <% Else %>
      <a href="?Page=<%=ThisCurrentPage-1%>&TxtKey=<%=TxtKey%>" class="blue">上一页</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      下一页
      <% Else %>
      <a href="?Page=<%=ThisCurrentPage+1%>&TxtKey=<%=TxtKey%>" class="blue">下一页</a>
      <% End If %>
    </td>
    <td align="center"><% If ThisCurrentPage=ThisPageCount Then %>
      最后一页
      <% Else %>
      <a href="?Page=<%=ThisPageCount%>&TxtKey=<%=TxtKey%>" class="blue">最后一页</a>
      <% End If %>
    </td>
  </tr>
</table>
<% 
	else
	response.Write"<script> alert('使用用户不准确，请返回'); history.back(); </script>"
	response.End()
	end if
 %>
</body>
</html>
<% call closeConn() %>
