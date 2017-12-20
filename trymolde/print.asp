<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../css/print.css" rel="stylesheet" type="text/css" />
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT>
</OBJECT>
<% 
Function getFormatDateString(dtNumber)
 Dim strDate
 If isNumeric(dtNumber) Then
  If dtNumber<10 Then
   dtNumber = "0" & dtNumber
  End If
  strDate = dtNumber
 Else
  strDate = "" 
 End if
 getFormatDateString = strDate
End Function
dastr=year(now)&"-"&getFormatDateString(month(now))&"-"&getFormatDateString(day(now))
 %>
</head>
 <style media="print">
    .noprint { display: none }
  </style>
<body style="font-size:10px;">
<table width="1050" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#000000" id="mytable">
  <tr style="font-weight:bold;">    
    <td height="60" colspan="16"><table width="1050" height="60" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="4" align="center" style="font-size:24px; font-weight:bold; font-family:'黑体'">试模联络单</td>
      </tr>
      <tr>
        <td align="left"><%= date() %>&nbsp;
          <form id="form1" name="form1" method="post" action="">
            <label>
            请输入纳期
            <input name="TxtKey" type="text" class="noprint" id="TxtKey" size="12" />
            </label>
            <label></label>
            <label class="noprint"> 例：2012-06-07 请输入机台</label>
            <label>
            <input name="equip_name" type="text" id="equip_name" size="12" />
            </label>
            <label class="noprint"> </label>
            <input type="submit" name="Submit" value="查 询" class="noprint" />
          </form>          </td>
        <td width="200" align="left">承认：</td>
      </tr>
    </table></td>
  </tr>
  <tr style="font-weight:bold;"> 
  <td>ID</td>   
    <td>类型</td>
	<td>型/品番</td>
    <td>机台</td>
	<td>辅助设备</td>  
    <td>穴数</td>
    <td>材料编号</td>
    <td>材料</td>
	<td>原料</td>
	<td>粉料</td>
    <td>数量</td>
    <td>PO</td>
	<td>旧品处理</td>
    <td>纳期</td>
    <td>担当</td>
    <td width="200">备注</td>
  </tr>
  <%
'定义查询关键字变量，从URL或表单中获取查询的关键字TxtKey，并组合成模糊查询SQL语句
Dim TxtKey
TxtKey=Request("TxtKey")
equip=request.Form("equip_name")
If TxtKey="" and equip="" Then
sql="select * from trymolde Where 纳期 ='"&dastr&"'"
else
sql="select * from trymolde Where 1=1"
if TxtKey<>"" then
sql=sql&" and 纳期 like '%"&TxtKey&"%'"
end if
if equip<>"" then
sql=sql&" and 生准要望机台='"& equip &"'"
end if
End If
sql=sql&" Order By 模具担当 DESC"
Call CreateRs(rs,sql,1,1)
		if rs.eof or rs.bof then
			response.write("没有试摸")
			response.End()
		end if	
i=1
Do While Not Rs.EOF
%>
  <tr>
  	<td><%= i %></td>
    <td><%= rs("试模类型") %> | <%= rs("类型") %>| <%= rs("mold_type") %></td>
	<td width="60"><%= rs("品番") %></td>
    <td width="60"><%= rs("生准要望机台") %></td>
	 <td><%= rs("fzsb") %></td>     
    <td width="40"><%= rs("穴数") %></td>
    <td width="80"><%= rs("材料编号") %></td>
    <td width="80"><%= rs("材料") %></td>
	<td><%= rs("用量") %>KG</td>
	<td width="60"><%= rs("用料F") %>KG</td>
    <td width="40"><%= rs("试作数量") %></td>
    <td><%= rs("PO") %></td>
	<td><%= rs("PO处理") %></td>
    <td><%= right( rs("纳期"),5) %></td>
    <td><%= rs("模具担当") %></td>
    <td><%= rs("备注") %></td>
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
<strong>今日试模点数：<%= i - 1  %></strong>
</body>
</html>
<%
Call CloseRs(rs)
call closeConn()
%>



