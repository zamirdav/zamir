<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>foto_add</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="initf.inc"-->
<%
'Response.Write "foto_add<BR>"
'Response.Write "davz<BR>"
'Response.Write "2021.10.21<BR>"



sub FOTO_FORMA(ID)
    Session("ID")=ID
'   Response.Write "ID=" & ID & "<BR>"
   Response.Write "<FORM method='post'  encType='multipart/form-data' action='foto_save.asp?ID=" & ID & "'>"
   Response.Write "<INPUT type='File'   name='File1'>"
   Response.Write "<INPUT type='Submit' value='посмотреть'>"
   Response.Write "</FORM>" 
end sub





ID = Request.QueryString("ID")
if ID&"~" = "~" then 
   Response.Write "нету параметра.<BR>"
else
'   call LICO_VID_ODIN(ID)
   call FOTO_FORMA(ID)
end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="VVO-form.asp?T=2"><div align=center><font color=Red> Выход без записи </div></A><BR>
