<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>foto_add</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="clsUpload.asp"-->
<%
'Response.Write "foto_add<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.21<BR>"


Dim Upload
Dim Folder
Dim FileName

   Folder = Server.MapPath("Uploads") & "\"
   Response.Write "Folder = " & Folder & "<BR>"

   Set Upload = New clsUpload

   FileName=Upload("File1").FileName
   Folder="C:\TEMP\1\"
   Response.Write "Сохраняем в = " & Folder & "<BR>"
   Response.Write "файл = " & FileName & "<BR>"

   Upload("File1").SaveAs Folder & FileName

   Set Upload = Nothing

    ID=Session("ID")

Response.Write "ID=" & ID & "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
