<% 
Dim DBPath,ConnStr,Conn,Rs,Sql
	ConnStr="PROVIDER=SQLOLEDB;DATA SOURCE=(local);database=hudsonwwwroot;uid=sa;pwd="
	Set Conn=Server.CreateObject("ADODB.Connection")
	Conn.Open ConnStr

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
