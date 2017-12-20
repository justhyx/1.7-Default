<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>试模联络单管理系统---打印</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT>
</OBJECT>
</head>
 <style media="print">

    .noprint { display: none }

  </style>
<body>
<table width="1050" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000" id="mytable">
  <tr style="font-weight:bold;">    
    <td height="60" colspan="13"><table width="1050" height="60" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2" align="center" style="font-size:24px; font-weight:bold; font-family:'黑体'">试模联络单</td>
      </tr>
      <tr>
        <td align="left"><%= date() %>&nbsp;
		<label class="noprint">
		          <form id="form1" name="form1" method="post" action="">
            <label>
            <input name="TxtKey" type="text" id="TxtKey" class="noprint" />
            </label>
            <label>
            <input type="submit" name="Submit" value="查 询" class="noprint" />
            </label>
            <label class="noprint">输入日期查询 例：2012-6-7 注：输入填单日期</label>
          </form>
		</label>		</td>
        <td width="200" align="left">承认：</td>
      </tr>
    </table></td>
  </tr>
  <tr style="font-weight:bold;">    
    <td width="90">型/品番</td>
    <td width="40">机台</td>
    <td width="60">OK机台</td>
    <td width="40">穴数</td>
    <td width="70">材料编号</td>
    <td width="40">材料</td>
    <td width="40">数量</td>
    <td width="40">PO</td>
    <td width="40">纳期</td>
    <td width="60"><a href="print.asp">担当</a></td>
    <td width="60">备注</td>
    <td width="60">时间安排</td>
    <td width="60">生管说明</td>
  </tr>
   <%  
'设置每页显示多少条记录的分页常量ThisPageSize
Const ThisPageSize=999
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
	call CreateRs(rs,"select * from trymolde where addtime='"&date()&"' and okfolg='"&int(0)&"'Order By 材料编号 DESC ",1,1)
		if rs.eof or rs.bof then
			response.End()
		end if
Else
	call CreateRs(rs,"select * from trymolde Where addtime like '%"&TxtKey&"%' Order By 材料编号 DESC",1,1)
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
    <td><%= rs("品番") %> | <%= rs("类型") %></td>
    <td><%= rs("生准要望机台") %></td>
    <td><%= rs("生管确认机台") %></td>
    <td><%= rs("穴数") %></td>
    <td><%= rs("材料编号") %></td>
    <td><%= rs("材料") %></td>
    <td><%= rs("试作数量") %></td>
    <td><%= rs("PO") %></td>
    <td><%= rs("纳期") %></td>
    <td><%= rs("模具担当") %></td>
    <td><%= rs("备注") %></td>
    <td><%= rs("安排时间") %>&nbsp;</td>
    <td><%= rs("生管备注") %>&nbsp;</td>
  </tr>
  <%
	i=i+1
	Rs.MoveNext
Loop	
%>
</table>
   <center class=noprint> <input onclick="document.all.WebBrowser.ExecWB(8,1)" type="button" value="页面设置">
&nbsp;<input onclick="document.all.WebBrowser.ExecWB(7,1)" type="button" value="打印预览">&nbsp;<input onclick="document.all.WebBrowser.ExecWB(6,1)" type="button" value="打印">
 </center>
 <strong>今日试模点数：<%= i %> </strong>
</body>
</html>
<%
Call CloseRs(rs)
call closeConn()
%>
