<%
Dim Conn,ConnStr
Set Conn=Server.CreateObject("Adodb.Connection")
ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(""&SysRootDir&""&SiteDataPath&"")
Conn.open ConnStr
If err Then
   err.clear
   Set Conn = Nothing
   Response.Write "���ݿ����Ӵ�������/Include/Const.asp�ļ�!"
   Response.End
End If

sub CreateRs(rs,sql,n1,n2)
	set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn,n1,n2
end sub

sub closeRs(rs)
	rs.close
	set rs=nothing
end sub

sub closeConn()
	conn.close
	set conn=nothing
end sub
%>
<!--#include file="Function.asp" -->