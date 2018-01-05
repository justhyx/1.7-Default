<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
'dim conn,str
'	set conn=server.CreateObject("adodb.connection")
'		str="provider=microsoft.Jet.OLEDB.4.0;data source="& server.MapPath("db/hudsonwwwroot.mdb")
'		conn.open str
'		
'sub CreateRs(rs,sql,n1,n2)
'	set rs=server.CreateObject("adodb.recordset")
'		rs.open sql,conn,n1,n2
'end sub
'
'sub closeRs(rs)
'	rs.close
'	set rs=nothing
'end sub
'
'sub closeConn()
'	conn.close
'	set conn=nothing
'end sub
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>和诚锴诚事务管理系统</title>
<link href="css/templatemo_style.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div id="templatemo_wrapper">
  <div id="templatemo_main">
	  <div id="templatemo_content">
	    <table width="950" height="490" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="150" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="login.asp">系统登录</a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="companycontent/companylist.asp">HUDSON规定标准</a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="mainger/QA_imageDown.asp">平板电脑通道</a></td>
              </tr>
              <tr>
                <td >&nbsp;</td>
              </tr>
              <tr>
                <td><a href="zc_condition/piclist.asp">成形条件</a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="PicFiles.asp">图档下载</a>&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="trymolde/addtrymolde.asp?event=add">试模联络单</a> | <a href="trymolde/print.asp">打印 </a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="softdown.asp">软件下载</a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="trymold_up/Upload_Photo.asp">试作确认表</a></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><a href="trymold_dh_up/Upload_Photo.asp">部品打合资料</a></td>
              </tr>
            </table></td>
            <td><table width="800" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="405" height="240" valign="top"><iframe id="frmleft" name="frmleft" width="400" height="230" src="meetdit.asp" 
      frameborder="0" allowtransparency="allowTransparency"></iframe></td>
                <td width="405" height="240" valign="top"><iframe id=frmleft name=frmleft width="400" height="240" src="audit.asp" 
      frameborder=0 allowtransparency></iframe></td>
              </tr>
              <tr>
                <td height="240" colspan="2" valign="top"><marquee onmouseover="this.stop()" onmouseout="this.start()" direction="up" scrollamount="2" height="230px">
               
</marquee></td>
              </tr>
            </table></td>
          </tr>
        </table>
	  </div>
	  <!-- END of content -->
		</div> 
	<!-- END of templatemo_main -->
    <div id="templatemo_footer"></div>
</div> <!-- END of templatemo_wrapper -->
</body>
</html>







