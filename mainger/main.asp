<% 
'	if session("UserName")="" then 
'	response.Redirect("../login.asp")
'	response.End()
'	end if
 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML 
xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>�ͳ��ǳ��������ϵͳ</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="images/style.css" type=text/css rel=stylesheet>
<STYLE>.main_left {
	TABLE-LAYOUT: auto; BACKGROUND: url(images/left_bg.gif)
}
.main_left_top {
	BACKGROUND: url(images/left_menu_bg.gif); PADDING-TOP: 5px
}
.main_left_title {
	PADDING-LEFT: 15px; FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #fff; TEXT-ALIGN: left
}
.left_iframe {
	BACKGROUND: none transparent scroll repeat 0% 0%; VISIBILITY: inherit; WIDTH: 180px; HEIGHT: 92%
}
.main_iframe {
	Z-INDEX: 1; VISIBILITY: inherit; WIDTH: 100%; HEIGHT: 92%
}
TABLE {
	FONT-SIZE: 12px; FONT-FAMILY: tahoma, ����, fantasy
}
TD {
	FONT-SIZE: 12px; FONT-FAMILY: tahoma, ����, fantasy
}
</STYLE>

<SCRIPT language=javaScript src="images/admin.js" 
type=text/javascript></SCRIPT>
<script language="javascript" src="../Script/Admin.js"></script>
<SCRIPT>
var status = 1;
var Menus = new DvMenuCls;
document.onclick=Menus.Clear;
function switchSysBar(){
     if (1 == window.status){
		  window.status = 0;
          switchPoint.innerHTML = '<img src="images/left.gif">';
          document.all("frmTitle").style.display="none"
     }
     else{
		  window.status = 1;
          switchPoint.innerHTML = '<img src="images/right.gif">';
          document.all("frmTitle").style.display=""
     }
}
</SCRIPT>

<META content="MSHTML 6.00.6000.16640" name=GENERATOR></HEAD>
<BODY style="MARGIN: 0px"><!--��������-->
<DIV class=top_table>
<DIV class=top_table_leftbg>

<DIV class=system_logo><IMG src="images/logo_up.gif"></DIV>
<DIV class=menu>
<UL>
  <LI id=menu_0 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">��������</a>
    <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>
    <LI><a href="../usecar/UseIndex.asp" target="frmright">�ó�����</a></LI>
	<% if session("Power")>999 then	 %>
    <LI><a href="../usecar/UseList.asp" target="frmright">�ó��鿴</a></LI>
	<% else %>
	<LI><a href="../usecar/approved.asp" target="frmright">�ó��鿴</a></LI>
	<% end if %>
	<li> <a href="../meeting/meeting.asp" target="frmright">������ʹ������</a> </LI>
	<li> <a href="../meeting/meetinglist.asp" target="frmright">������ʹ�ò鿴</a> </LI>
	<li> <a href="../overtime/ot_add.asp" target="frmright">Ԥ���ڼӰ�����</a> </LI>
	<li> <a href="../overtime/list_ot.asp" target="frmright">�Ӱ�鿴</a> </LI>
	<% if session("username")="��ӳϼ" then %>
	<li> <a href="../overtime/list_ot_m.asp" target="frmright">����Ӱ�ͳ��</a> </LI>
	<%end if%>
	<li> <a href="../overtime/other_add.asp" target="frmright">Ԥ����Ӱ�����</a> </LI>
	<% if session("user_l")="21" or session("user_l")="11" then %>
	<li> <a href="../wifi/wifipassword.asp" target="frmright">wifi����</a> </LI>	
	<% end if %>
	<li> <a href="../ITevent/eventlook.asp" target="frmright">IT��������</a> </LI>
	<% if session("user_l")="22" or session("user_l")="11" then %>
	<li> <a href="../wifi/computer02.asp" target="frmright">���Թ���</a> </LI>
	<% end if %>
	<% if session("user_l")="11" then %>
	<li> <a href="../wifi/eamiladd.asp" target="frmright">�������</a> </LI>
	<% end if %>
	<li> <a href="../EPevent/eventlook.asp" target="frmright">�豸ά������</a> </LI>
   </UL></DIV>
   
   <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <LI id=menu_1 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">��׼����</a>
    <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>
    <LI><a href="../trymolde/trymoldelist.asp" target="frmright">��ģ���絥</a></LI>
    <LI><a href="../trymolde/arrange.asp" target="frmright">������ģ����</a></LI>
	 <LI><a href="../trymolde/statements.asp" target="frmright">ģ��״̬</a></LI>
	 <LI><a href="../trymold_up/Upload_Photo.asp" target="frmright">��ģȷ�ϱ�</a></LI>
         <LI><a href="../trymold_dh_up/Upload_Photo.asp" target="frmright">�������</a></LI>
	 <% if session("Power")<99 then %>
	 <LI><a href="http://192.168.1.7:81/CrushedMaterialManager/CrushedMaterial_List.aspx" target="frmright">��ģƷ����</a></LI>   <% end if %> 
	 <LI><a href="http://192.168.1.7:81/CrushedMaterialManager/CrushedMaterial_Approved.aspx 
" target="frmright">������鿴</a></LI>
	 </UL></DIV>
   
  <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  
  <LI id=menu_2 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">ͼ������</a>
  <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>  
  <li><a href="../PicUP/Upload_Photo.asp" target="frmright">ͼ�����</a></LI>
  <li><a href="../PicUP/Piclist.asp" target="frmright">ͼ��ȷ��</a></LI>  
  </UL></DIV>
  <!--<DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <LI id=menu_4 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">����������</a> 
    <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>
    <li> <a href="../meeting/meeting.asp" target="frmright">������ʹ������</a> </LI>
	    <li> <a href="../meeting/meetinglist.asp" target="frmright">������ʹ�ò鿴</a> </LI>          
	</UL></DIV>-->
  <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <LI id=menu_5 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">Ʒ�����</a> 
    <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>
    <li><a href="QA_image.asp" target="frmright">�ϴ������׼</a></LI>
		<li><a href="QA_imageDown.asp" target="frmright">�鿴�����׼</a></LI>
		<li><a href="http://192.168.1.7:81/QCCheckManager/EngineeringTest_Show.aspx?u_id=<%= session("user_id")  %>" target="frmright">�������</a></LI>
		<li><a href="http://192.168.1.7:81/QCCheckManager/EngineeringTest_Add_Tool.aspx" target="frmright">�������ⶨ�߹���</a></LI>
		<li><a href="http://192.168.1.199:81/MaterialControlManager/BarcodRecordsExport.aspx" target="view_window">ɨ���¼����</a></LI>
		<li><a href="http://192.168.1.146:81/QCCheckManager/EngineeringTest_Show.aspx?UserName=UserID" target="view_window">�������ݵ���</a></LI>
		<!--<li><a href="#" target="frmright">XXX</a></LI>-->		
		</UL></DIV>
  <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <LI id=menu_7 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">�ƶȹ���</A> 
    <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>
    <li><a href="../companycontent/addparent.asp" target="frmright">�涨�ϴ�</a></LI>
	<!--<li><a href="#" target="frmright">XXX</a></LI>
	<li><a href="#" target="frmright">XXX</a></LI>-->
	</UL></DIV>
  <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <LI id=menu_8 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">�û�����</a> 
  <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>

		<li><a href="UserMangerEidet.asp?action=add" target="frmright">�û�������</a></LI>				
		<li><a href="UserManger.asp" target="frmright">�û��鿴</a></LI></UL></DIV>  
  <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <% if session("UserName")="���" then %>
  <LI id=menu_10 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="zj.asp" target="frmright">������־</a>
  <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <!--<UL>
    
          <li><a href="#" target="frmright">XXX</a></LI>
		  <li><a href="#" target="frmright">XXX</a></LI>		
		  <li><a href="#" target="frmright">XXX</a></LI>			  
		  <li><a href="#" target="frmright">XXX</a></LI>
		  <li><a href="#" target="frmright">XXX</a></LI>		
</UL>--></DIV>
 <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  <% end if %>  
  <LI id=menu_9 onmouseover=Menus.Show(this,0) onclick=getleftbar(this);><a href="#">���ι���</a>
  <DIV class=menu_childs onmouseout=Menus.Hide(0);>
  <UL>
   <LI><A href="../zc_condition/Upload_Photo.asp" target="frmright">������������</A></li>
   <!--<LI><A href="#" target=_blank>XXX</A></LI>   
   <LI><A href="#" target=_blank>XXX</A></LI>-->
				</UL></DIV>

  
    <DIV class=menu_div><IMG style="VERTICAL-ALIGN: bottom" 
  src="images/menu01_right.gif"></DIV></LI>
  
  </UL></DIV></DIV></DIV>
<DIV style="BACKGROUND: #337abb; HEIGHT: 24px"></DIV><!--�������ֽ���-->
<TABLE style="BACKGROUND: #337abb" height="92%" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD class=main_left id=frmTitle vAlign=top align=middle name="fmTitle">
      <TABLE class=main_left_top cellSpacing=0 cellPadding=0 width="100%" 
      border=0>
        <TBODY>
        <TR height=32>
          <TD vAlign=top></TD>
          <TD class=main_left_title id=leftmenu_title>���ÿ�ݹ���</TD>
          <TD vAlign=top align=right></TD></TR></TBODY></TABLE><IFRAME 
      class=left_iframe id=frmleft name=frmleft src="left.htm" 
      frameBorder=0 allowTransparency></IFRAME>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR height=32>
          <TD vAlign=top></TD>
          <TD vAlign=bottom align=middle></TD>
          <TD vAlign=top align=right></TD></TR></TBODY></TABLE></TD>
    <TD style="WIDTH: 10px" bgColor=#337abb>
      <TABLE height="100%" cellSpacing=0 cellPadding=0 border=0>
        <TBODY>
        <TR>
          <TD style="HEIGHT: 100%" onclick=switchSysBar()><SPAN class=navPoint 
            id=switchPoint title=�ر�/������><IMG 
            src="images/right.gif"></SPAN> </TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top width="100%" bgColor=#337abb>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" bgColor=#c4d8ed 
        border=0><TBODY>
        <TR height=32>
          <TD vAlign=top width=10 background=images/bg2.gif><IMG 
            alt="" src="images/teble_top_left.gif"></TD>
          <TD width=28 background=images/bg2.gif></TD>
          <TD background=images/bg2.gif><SPAN 
            style="FLOAT: left">�ͳ��ǳ��������ϵͳ</SPAN><SPAN id=dvbbsannounce 
            style="FONT-WEIGHT: bold; FLOAT: left; WIDTH: 300px; COLOR: #c00"></SPAN></TD>
          <TD style="COLOR: #135294; TEXT-ALIGN: right" 
          background=images/bg2.gif>��ӭ����<%= session("Department") %> &nbsp; <%= session("UserName") %> | <a href="javascript:history.go(-1);"> ����</a> | <A 
            href="main.asp" target=_top>����̨</A> |<a href="out.asp" target=_top> �˳�</A> </TD>
          <TD vAlign=top align=right width=28 
            background=images/bg2.gif><IMG alt="" 
            src="images/teble_top_right.gif"></TD>
          <TD align=right width=16 bgColor=#337abb></TD></TR></TBODY></TABLE><IFRAME 
      class=main_iframe id=frmright name=frmright 
      src="syscome.asp" frameBorder=0 scrolling=yes></IFRAME>
      <TABLE style="BACKGROUND: #c4d8ed" cellSpacing=0 cellPadding=0 
      width="100%" border=0>
        <TBODY>
        <TR>
          <TD><IMG height=6 alt="" src="images/teble_bottom_left.gif" 
            width=5></TD>
          <TD align=right><IMG height=6 alt="" 
            src="images/teble_bottom_right.gif" width=5></TD>
          <TD align=right width=16 
  bgColor=#337abb></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<DIV id=dvbbsannounce_true style="DISPLAY: none">
</DIV>
<SCRIPT language=JavaScript>
<!--
document.getElementById("dvbbsannounce").innerHTML = document.getElementById("dvbbsannounce_true").innerHTML;
//-->
</SCRIPT>
</BODY></HTML>







