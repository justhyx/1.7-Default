<!--#include file="../connet/conn.asp" -->
<% 
dim IT_ID,IT_Department,IT_UserName,IT_Time,IT_Tel,IT_Approve,IT_Eventone,IT_Eventtwo,IT_Eventthree,IT_Eventfour,IT_Eventfive,IT_Eventsix,IT_Eventseven,IT_Eventeight,IT_Eventnight,IT_Eventten,IT_Eventeleven,IT_Eventtwelve,IT_Eventthirteen,IT_Eventfourteen,IT_Eventwhy,IT_Eventdetail,IT_Eventresult
randomize
ranNum=int(900*rnd)+100
IT_ID=year(now)&month(now)&day(now)&timer()*100&ranNum



 call closeConn() %>