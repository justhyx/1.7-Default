<%
Sub DoDelslhtml(htmlname)
	On Error Resume Next
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	servermap=server.MapPath("..")
	servermap=servermap&"\"&htmlname
	FSO.DeleteFile(servermap)
	Set FSO = Nothing
End Sub

Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

Function ReplaceConstChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceConstChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceConstChar = tempChar
End Function

function StrLen(Str)
  if Str="" or isnull(Str) then
    StrLen=0
    exit function
  else
    dim regex
    set regex=new regexp
    regEx.Pattern ="[^\x00-\xff]"
    regex.Global =true
    Str=regEx.replace(Str,"^^")
    set regex=nothing
    StrLen=len(Str)
  end if
end function

function StrLeft(Str,StrLen)
  dim L,T,I,C
  if Str="" then
    StrLeft=""
    exit function
  end if
  Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
  L=Len(Str)
  T=0
  for i=1 to L
    C=Abs(AscW(Mid(Str,i,1)))
    if C>255 then
      T=T+2
    else
      T=T+1
    end if
    if T>=StrLen then
      StrLeft=Left(Str,i) & "…"
      exit for
    else
      StrLeft=Str
    end if
  next
  StrLeft=Replace(Replace(Replace(replace(StrLeft," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

function StrReplace(Str)
  if Str="" or isnull(Str) then
    StrReplace=""
    exit function
  else
    StrReplace=replace(str," ","&nbsp;")
    StrReplace=replace(StrReplace,chr(13),"&lt;br&gt;")
    StrReplace=replace(StrReplace,"<","&lt;")
    StrReplace=replace(StrReplace,">","&gt;")
  end if
end function

function ReStrReplace(Str)
  if Str="" or isnull(Str) then
    ReStrReplace=""
    exit function
  else
    ReStrReplace=replace(Str,"&nbsp;"," ")
    ReStrReplace=replace(ReStrReplace,"<br />",chr(13))
    ReStrReplace=replace(ReStrReplace,"&lt;br&gt;",chr(13))
    ReStrReplace=replace(ReStrReplace,"&lt;","<")
    ReStrReplace=replace(ReStrReplace,"&gt;",">")
  end if
end function

function HtmlStrReplace(Str)
  if Str="" or isnull(Str) then
    HtmlStrReplace=""
    exit function
  else
    HtmlStrReplace=replace(Str,"&lt;br&gt;","<br />")
  end if
end function

function ViewNoRight(GroupID,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from WrtxcnHTML_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  ViewNoRight=true
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then
	    ViewNoRight=false
	  end if
    case "="
      if not session("GroupLevel") = GroupLevel then
	    ViewNoRight=false
      end if
  end select
end function

Function GetUrl()
  GetUrl="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
  If Request.ServerVariables("QUERY_STRING")<>"" Then GetURL=GetUrl&"?"& Request.ServerVariables("QUERY_STRING")
End Function

function HtmlSmallPic(GroupID,PicPath,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from WrtxcnHTML_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  HtmlSmallPic=PicPath
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then HtmlSmallPic="../Images/NoRight.jpg"
    case "="
      if not session("GroupLevel") = GroupLevel then HtmlSmallPic="../Images/NoRight.jpg"
  end select
  if HtmlSmallPic="" or isnull(HtmlSmallPic) then HtmlSmallPic="../Images/NoPicture.jpg"
end function

function IsValidMemName(memname)
  dim i, c
  IsValidMemName = true
  if not (3<=len(memname) and len(memname)<=16) then
    IsValidMemName = false
    exit function
  end if
  for i = 1 to Len(memname)
    c = Mid(memname, i, 1)
    if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-", c) <= 0 and not IsNumeric(c) then
      IsValidMemName = false
      exit function
    end if
  next
end function

function IsValidEmail(email)
dim names, name, i, c
  IsValidEmail = true
  names = Split(email, "@")
  if UBound(names) <> 1 then
    IsValidEmail = false
    exit function
  end if
  for each name in names
	if Len(name) <= 0 then
	  IsValidEmail = false
      exit function
    end if
    for i = 1 to Len(name)
      c = Mid(name, i, 1)
      if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-.", c) <= 0 and not IsNumeric(c) then
        IsValidEmail = false
        exit function
      end if
	next
	if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
    end if
  next
  if InStr(names(1), ".") <= 0 then
    IsValidEmail = false
    exit function
  end if
  i = Len(names(1)) - InStrRev(names(1), ".")
  if i <> 2 and i <> 3 then
    IsValidEmail = false
    exit function
  end if
  if InStr(email, "..") > 0 then
    IsValidEmail = false
  end if
end function

Function FormatDate(DateAndTime, Format)
  On Error Resume Next
  Dim yy,y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(Format) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  yy = CStr(Year(DateAndTime))
  y = Mid(CStr(Year(DateAndTime)),3)
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
  Select Case Format
  Case "1"
    strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
    strDateTime = yy & m & d & h & mi & s
  Case "3"
    strDateTime = yy & m & d & h & mi
  Case "4"
    strDateTime = yy & "年" & m & "月" & d & "日"
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "月" & d & "日"
  Case "8"
    strDateTime = y & "年" & m & "月"
  Case "9"
    strDateTime = y & "-" & m
  Case "10"
    strDateTime = y & "/" & m
  Case "11"
    strDateTime = y & "-" & m & "-" & d
  Case "12"
    strDateTime = y & "/" & m & "/" & d
  Case "13"
    strDateTime = yy & "." & m & "." & d
  Case "14"
    strDateTime = yy & "-" & m & "-" & d
  Case Else
    strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function

function WriteMsg(Message)
  response.write "<script language='JavaScript'>alert('"&Message&"');" & "history.back()" & "</script>"
end function

public SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,Video,IcpNumber,MesViewFlag
sub SiteInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from WrtxcnHTML_Site"
  rs.open sql,conn,1,1
  SiteTitle=rs("SiteTitle")
  Keywords=rs("Keywords")
  Descriptions=rs("Descriptions")
  SiteUrl=rs("SiteUrl")
  ComName=rs("ComName")
  Address=rs("Address")
  ZipCode=rs("ZipCode")
  Telephone=rs("Telephone")
  Fax=rs("Fax")
  Email=rs("Email")
  Video=rs("Video")
  IcpNumber=rs("IcpNumber")
  MesViewFlag=rs("MesViewFlag")
  rs.close
  set rs=nothing
end Sub

  Function CheckStr(Strer,Num)
        Dim Shield,w
	    If Strer = "" Or IsNull(Strer) Then Exit Function
	    Select Case Num
		  Case 1
	        If IsNumeric(Strer) = 0 Then
	          Response.Write "操作错误"
		      Response.End
	        End If
			Strer = Int(Strer)
		End Select
	    CheckStr = Strer
   End Function
%>