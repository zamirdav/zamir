<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>foto_save</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="clsUpload.asp"-->
<!--#INCLUDE FILE="initf.inc"-->
<%
'Response.Write "foto_save<BR>"
'Response.Write "zamir<BR>"
'Response.Write "2022.10.10<BR>"

'Dim Upload
'Dim Folder
'Dim FileName

   Set Upload = New clsUpload
   ADRES = Request.QueryString("ID") 
   FIL_SIZE = Upload.Fields("File1").Length
   F_NAME = Upload.Fields("File1").FileName
   f_Path = Upload.Fields("File1").FilePath
   Path = Server.MapPath(F_NAME)
'   Response.Write" ADRES "& ADRES&"<BR>"
'   Response.Write "ID=" &ADRES& "<HR>"   
'   Response.Write "Path=" & Path& "<HR>" 
'   Response.Write "F_NAME=" & F_NAME & "<HR>"
'   Response.Write "FIL_SIZE=" & FIL_SIZE & "<HR>"
'   Response.Write "FilePath=" & FilePath & "<HR>"
   Response.Write "<A href='vvo-form.asp?T=1'><div align=center><font color=Red> Выход без записи <IMG src=назад.gif border=0 alt=Выход></div></A><BR>"
   PATH ="/foto_gala/"
   FIL_NAME =PATH & F_NAME
   Response.Write "<A href='add_f.asp?FIL_NAME="&FIL_NAME & "&F_Name="&F_NAME& "&FIL_SIZE="&FIL_SIZE & "&ADRES="&ADRES&"'><div align=bottom> Записать </A><BR>"
   Response.Write "<EMBED src='" &FIL_NAME&"'  height=540 width =1400></EMBED>"
   Set Upload = Nothing




%>
