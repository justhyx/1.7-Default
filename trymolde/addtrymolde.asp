<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../connet/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>试模联络单管理</title>
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
</head>
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
<%
ev=request.QueryString("event")
if ev="add" then
actvalue="addcentent"
 keypress="添  加"
else
actvalue="exchange"
 keypress="修  改"
end if


%>
</script>
<body>
<table width="1050" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
	<tr>
		<td align="left">
<form id="form1" name="form1" method="post" action="chk.asp?action=addnumber">
  <table width="950" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td height="50" colspan="4" align="center" bgcolor="#FFFFFF" style="font-size:24px; font-weight:bold; font-family:'黑体'">试模联络单</td>
    </tr>
    <tr>
      <td width="137" height="20" align="right" bgcolor="#FFFFFF" >型/品番
      <label></label></td>
      <td width="320" align="left" bgcolor="#FFFFFF" ><input name="t2" type="text" id="t2" />
      <input type="submit" name="Submit2" value="获取部品信息" /></td>
      <td width="301" align="left" bgcolor="#FFFFFF" ><a href="search.asp">查询试模联络单</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://192.168.1.7/trymolde/print.asp">试模联络单打印</a></td>
      <td align="right" bgcolor="#FFFFFF" ><p>&nbsp;</p>        </td>
    </tr>
  </table> 
</form>
<form id="form2" name="form2" method="post" action="chk.asp?action=<%=actvalue%>">
     <%
	 dim at,bt,ct
	 'ndd=Request.QueryString("numberid")
	 ndd=Cint(Request.QueryString("numberid"))
	if request.QueryString("event")="add" then
	 call createRs(rsm,"select *from trymolde where id='"& ndd &"'",1,1)
	 if not rsm.eof then
	 numberidadd=rsm("品番")
	 end if
	 call closeRs(rsm)
	 if numberidadd="" then
	 	at=""
		bt=""
		cl=""
		dt=""
		et=""	
	else
		call CreateRs(rsa,"select * from compose where 部品番号='"& numberidadd &"'",1,1)
			if rsa.eof then
				at=""
				bt=""
				cl=""
				dt=""
				et=numberidadd
				nq=""
				szsl=""
				yl=""
				ylf=""
				posl=""
				bz=""
			else	
				at=rsa("穴数")
				bt=rsa("部品材料编号")
				cl=rsa("名称")
				dt=rsa("机位")
				et=rsa("部品番号")
				nq=rsa("纳期")
				szsl=rsa("试作数量")
				yl=rsa("用量")
				ylf=rsa("用料F")
				posl=rsa("PO")
				bz=rsa("备注")
				mjdd=rsa("模具担当")
				smlx=rsa("试模类型")
				lx=rsa("类型")
				poc=rsa("PO处理")
			end if
			call closeRs(rsa)
	end if
elseif request.QueryString("event")="change" then
			call CreateRs(rsa,"select * from trymolde where id='"& ndd &"'",1,1)
			if not rsa.eof then
				at=rsa("穴数")
				bt=rsa("材料编号")
				cl=rsa("材料")
				et=rsa("品番")
				nq=rsa("纳期")
				dt=rsa("生准要望机台")
				szsl=rsa("试作数量")
				yl=rsa("用量")
				ylf=rsa("用料F")
				posl=rsa("PO")
				bz=rsa("备注")
				mjdd=rsa("模具担当")
				smlx=rsa("试模类型")
				lx=rsa("类型")
				poc=rsa("PO处理")
				end if
	end if
 %>
  <table width="850" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td colspan="4" bgcolor="#FFFFFF"><hr width="100%" size="1" noshade="noshade" /></td>
    </tr> 
	</table>
  <table width="850" border="1" cellpadding="4" cellspacing="0" bordercolor="#030303" bgcolor="#CCCCCC">
    <tr>
      <td width="150" align="right" bgcolor="#FFFFFF">型/品番</td>
      <td width="300" colspan="2" bgcolor="#FFFFFF"><label></label>
      <input name="tx" type="text" id="tx" value="<%= et %>" disabled="true"/><input type="hidden" name="tryid" value="<%= ndd%>"></td>
      <td width="100" align="right" bgcolor="#FFFFFF">模具担当</td>
      <td width="150" bgcolor="#FFFFFF"><label>
      <select name="t1" id="t1">
	<option value="<%=mjdd%>" selected="selected"><%=mjdd%></option>
        <option value="张峰">张峰</option>
        <option value="庞宇">庞宇</option>
        <option value="高洁">高洁</option>
	<option value="姜波">姜波</option>
	<option value="罗彩文">罗彩文</option>
	<option value="龙麟">龙麟</option>
	<option value="牟彦澄">牟彦澄</option>
        <option value="张志祥">张志祥</option>
        <option value="张剑">张剑</option>
        <option value="陈英国">陈英国 </option>
        <option value="姚承">姚承</option>
	<option value="魏瑞红">魏瑞红</option>
        <option value="周学文">周学文</option>
	<option value="龚本灿">龚本灿</option>
	<option value="苏本方">苏本方</option>
	<option value="夏生军">夏生军</option>
	<option value="吴发明">吴发明</option>
              </select>
      </label></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">试模类型</td>
      <td bgcolor="#FFFFFF"><label>
        <select name="s1" id="s1">
		<option value="<%=smlx%>" selected="selected"><%=smlx%></option>
          <option value="新规">新规试模</option>
          <option value="修理">修理试模</option>
        </select>
      </label></td>
      <td bgcolor="#FFFFFF">模具类型
        <select name="mold_type" id="mold_type">
          <option value="量产">量产模具</option>
          <option value="贩卖">贩卖模具</option>
        </select>
        </td>
      <td align="right" bgcolor="#FFFFFF">生准要望机台</td>
      <td bgcolor="#FFFFFF"><input name="t3" type="text" id="t3" value="<%= dt %>" /></td>
    </tr>
    <tr>
	  <td align="right" bgcolor="#FFFFFF">穴数</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="t5" type="text" id="t5" value="<%= at %>" /></td>
      <td align="right" bgcolor="#FFFFFF">材料编号</td>
      <td bgcolor="#FFFFFF"><input name="t6" type="text" id="t6" value="<%= bt %>" /></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">材料</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="t7" type="text" id="t7" value="<%=cl %>" size="60" /></td>
      <td align="right" bgcolor="#FFFFFF">原料
        <input name="t82" type="text" id="t82" size="5" value="<%=yl%>"/>
        KG</td>
      <td align="left" bgcolor="#FFFFFF">粉碎料
        <input name="t822" type="text" id="t822" size="5" value="<%=ylf%>"/>
        KG</td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">纳期</td>
      <td colspan="2" bgcolor="#FFFFFF">
        <input name="t9" type="text" id="t9" value="<% =nq %>"onClick="WdatePicker()" style="hand();" />
        <img onclick="WdatePicker({el:$dp.$('t9')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle"></td>
      <td align="right" bgcolor="#FFFFFF">试作数量</td>
      <td bgcolor="#FFFFFF"><input name="t8" type="text" id="t8" value="<% =szsl%>"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">类型</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="t11" id="t11">
	  <option value="<%=lx%>" selected="selected"><%=lx%></option>
        <option value="内作">内作</option>
        <option value="外发">外发</option>
        <option value="移行">移行</option>
      </select>
      辅助设备
      <select name="fzsb" id="fzsb">
        <option value="模温">模温</option>
        <option value="热流道">热流道</option>
        <option value="模温 热流道">模温 热流道</option>
        <option value="0" selected="selected">请选择</option>
      </select></td>
      <td align="right" bgcolor="#FFFFFF">PO数量
        <input name="t12" type="text" id="t12" size="5" value="<%= posl%>"/></td>
      <td bgcolor="#FFFFFF">在库旧品处理
        <select name="POchk" id="POchk">
		  <option value="" selected="selected"></option>
		  <option value="<%=poc%>" selected="selected"><%=poc%></option>
          <option value="粉碎">保留</option>
          <option value="废弃">废弃</option>
              </select></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">备注</td>
      <td colspan="2" bgcolor="#FFFFFF"><textarea name="b1" rows="4" id="b1" value="<%=bz%>"></textarea></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="<%= keypress%>" />
        <label>
<input name="radiobutton" type="checkbox" id="radiobutton" value="radiobutton" checked="checked" />        
记忆</label>        </tr> 
</table>
</form>
</td>
</tr>
</table>	
  <hr align="center" width="1050" size="2" noshade="noshade" />
  <table width="1250" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#030303" bgcolor="#F1F1F1" id="mytable">
    <tr align="center" style="font-weight:bold;">
      <td>试模类型</td>
      <td>型/品番</td>
      <td>要望机台</td>
	  <td>辅助设备</td>
      <td>穴数</td>
      <td>材料编号</td>
      <td>材料</td>
      <td>试作数量</td>
      <td>原料</td>
      <td>粉碎料</td>
      <td>PO</td>
	  <td>在库旧品处理</td>
      <td>纳期</td>
      <td>模具担当</td>
      <td>备注</td>
      <td>日期</td>
      <td>操作</td>
    </tr>
<%
	Dim TxtKey
	TxtKey=Request("TxtKey")
	If TxtKey="" Then
		dress=Request.ServerVariables("REMOTE_ADDR")
		call CreateRs(rs,"select * from trymolde where okfolg='"&int(1)&"' and dress='"&dress&"' and addtime='"&date()&"' and (CONVERT(datetime, 纳期) >= DATEADD(day, - 1, GETDATE())) Order By id DESC ",1,1)
			if rs.eof or rs.bof then
				response.End()
			end if
	Else
		call CreateRs(rs,"select * from trymolde Where 品番 like '%"&TxtKey&"%' and (CONVERT(datetime, 纳期) > GETDATE()) Order By ID DESC",1,1)
			if rs.eof or rs.bof then
				response.Write("没有此记录")
				response.End()
			end if
	End If	
	i=0
	Do While Not Rs.EOF
	%>
    <tr>
      <td><%= rs("试模类型") %> | <%= rs("类型") %> | <%= rs("mold_type") %></td>
      <td><%= rs("品番") %>&nbsp;</td>
      <td><%= rs("生准要望机台") %>&nbsp;</td>
	  <td><%= rs("fzsb") %>&nbsp;</td>
      <td><%= rs("穴数") %>&nbsp;</td>
      <td><%= rs("材料编号") %>&nbsp;</td>
      <td><%= rs("材料") %>&nbsp;</td>
      <td><%= rs("试作数量") %>&nbsp;</td>
      <td><%= rs("用量") %>&nbsp;</td>
      <td><%= rs("用料F") %>&nbsp;</td>
      <td><%= rs("PO") %></td>
	  <td><%= rs("PO处理") %></td>
      <td>&nbsp;<%= rs("纳期") %></td>
      <td><%= rs("模具担当") %>&nbsp;</td>
      <td><%= rs("备注") %></td>
      <td><%= rs("addtime") %></td>
      <td><% if session("UserName")="庞宇" or session("UserName")="余艳军" then  %>
          <a href="chk.asp?id=<%= rs("id")%>&amp;action=modify&op=isok&ipdress=addtrymolde">承认</a> |
        <% end if %>
		 <a href="addtrymolde.asp?numberid=<%= rs("id")%>&amp;event=change" >更改</a> |
        <a href="chk.asp?id=<%= rs("id")%>&amp;action=modify&op=isno&ipdress=addtrymolde" onclick="return confirm('确定要取消这条记录吗？')">取消</a></td>
    </tr>
    <%
	i=i+1
	Rs.MoveNext
	Loop	
	%>
</table>
</body>
</html>
<%
call closeConn()
%>



