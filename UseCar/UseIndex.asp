<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../Images/Admin.css" rel="stylesheet" type="text/css" />
 <script language="JavaScript" type="text/javascript" src="../../My97DatePicker/WdatePicker.js"></script>
 <link href="../mainger/Images/Admin.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form id="form1" name="form1" method="post" action="check.asp?action=add">
  <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F1F1F1">
    <tr align="center">
      <td colspan="4" style="font-weight:bold">用车申请表</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="129" align="right">申请部门</td>
      <td width="230"><label>
        <input name="CarDepartment" type="text" id="CarDepartment" value="<%= session("Department")%>" />
      </label></td>
      <td width="108" align="right">申请人： </td>
      <td width="196"><label>
        <input name="UsePeople" type="text" id="UsePeople" value="<%= session("UserName")%>" />
      </label>        
        &nbsp;*必填</td>
    </tr>
    <tr>
      <td align="right">预定起始日期</td>
      <td><input name="StartTime" type="text" id="StartTime" onClick="WdatePicker()" style="hand();" />
        <img onclick="WdatePicker({el:$dp.$('StartTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle"></td>
      <td align="right">至</td>
      <td><input name="OverTime" type="text" id="OverTime" onClick="WdatePicker()" />
      <img onclick="WdatePicker({el:$dp.$('OverTime')})" src="../My97DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">目的地</td>
      <td><input name="Destination" type="text" id="Destination" /></td>
      <td align="right">预计用车时间</td>
      <td><input name="UseTime" type="text" id="UseTime" size="8" />
      小时</td>
    </tr>
    <tr >
      <td align="right">预计出车时间</td>
      <td><label>
<select name="hh" id="hh">
<option value=" ">请选择</option>
  <% for i=0 to 24 %>
  <option value="<%= i %>"><%= i %></option>
  <% next %>
</select>      
时
<select name="mm" id="mm">
  <option value="00:00">00</option>
  <option value="15:00">15</option>
  <option value="30:00">30</option>
  <option value="45:00">45</option>
</select>
      分　注：未提前3小时申请用车均视为紧急用车需要填写紧急理由</label></td>
      <td align="right">选择车型</td>
      <td><label>
      <select name="cartype" id="cartype">
        <option value="轿车">轿车</option>
        <option value="商务车">商务车</option>
		<option value="皮卡">皮卡</option>
        <option value="1.5吨货车">1.5吨货车</option>
        <option value="5吨货车_开顶">5吨货车_开顶</option>
        <option value="5吨货车_不开顶">5吨货车_不开顶</option>
        <option value="9.8吨货车">9.8吨货车</option>
        <option value="0" selected="selected">请选择</option>
      </select>
      </label></td>
    </tr>
    <tr >
      <td align="right">载人数</td>
      <td><label>
      <input name="peopleqty" type="text" id="peopleqty" />
      人</label></td>
      <td align="right">预计出货卡板数</td>
      <td><label>
        <input name="pl" type="text" id="pl" size="8" />
        卡板
      </label></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">事 由</td>
      <td colspan="3"><label>
        <textarea name="Reason" cols="40" rows="4" id="Reason"></textarea>
      *必须填写</label></td>
    </tr>
    <tr>
      <td align="right">* 临时紧急用车事由 </td>
      <td colspan="3"><textarea name="Urgency" cols="40" rows="4" id="Urgency"></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right">* 改善措施 </td>
      <td colspan="3"><textarea name="Improve" cols="40" rows="4" id="Improve"></textarea></td>
    </tr>
	<% if session("Power")<9 then %>
    <tr>
      <td rowspan="2" align="right">用车审批签署</td>
      <td>*董事长/总经理/副总经理临时紧急用车签署 </td>
      <td colspan="2">部门经理</td>
    </tr>
    <tr>
      <td>
	  <% if session("Power")<5 then %>
	  <label>
        <select name="UrgencyAgree" id="UrgencyAgree">
          <option value="1" selected="selected">同意</option>
          <option value="0">不同意</option>
        </select>
      </label>
	  <% end if %>	  </td>
      <td colspan="2"><select name="Agree" id="Agree">
        <option value="1" selected="selected">同意</option>
        <option value="0">不同意</option>
      </select></td>
    </tr>
	<% end if %>
  </table>
  <table width="100%" border="0" align="center" cellpadding="12" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>备注：1：各部门用厂车必须提前计划安排,需提前填写《用车申请单》，审批签署好后递交至 <br />
总务部，方可安排调达车辆。 <br />
提前交单时间期限：a) 明天上午用车，需于前一天下班(即PM5:30)前递交. <br />
b) 当天下午用车,需提前3小时前递交. <br />
c) 当天晚上用车,需于当天下午3:00前递交. <br />
2：正常申请用车,由部门经理签署生效. <br />
3: 临时紧急用车须由部门经理和(董事长/总经理/副总经理)任其一人共同签署生效. <br />
临时紧急用车以上“ * ” 栏必须填写方有效. <br />
4：周六周日及节假日用车或借车,必须在星期五PM5:30前递交《用车申请单》,并且须由 <br />
部门经理和(董事长/总经理/副总经理)任其一人共同签署生效. </td>
    </tr>
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr align="center">
      <td><label>
        <input name="action" type="hidden" id="action" value="add" />
        <input type="submit" name="Submit" value="提 交" />
      </label></td>
    </tr>
  </table>
  <br />
</form>
</body>
</html>
