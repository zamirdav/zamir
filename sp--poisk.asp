
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>поиск ALL</title>
</HEAD>

<!--#INCLUDE FILE="sp--poisk.inc"--> 
<%
' 2020.08.26 переставил реквизиты местами
'2021.08.12 Or F1="" поставил для статистики строка 40 для внутреннего вложения 
'
'
Tim_start=Time
'Response.Write "Довлетбаев З.С.<BR>"
'Response.Write "2019.01.24<BR>"
'27.11.2020 добавил поиск по ОСК О3


Response.Write "<DIV align=right><A href='pass.asp'>&nbsp;" & KOT_FIO & "&nbsp;</A></DIV>"  

T = Request.QueryString("T") 

if T="1" then
   if F1="" then 
      Response.Write "без фамилии искать не буду !<BR>"
      Break
   end if
'   Call POISK_ZAD("O2", 1, "07", "08", "09", "12", "OSK_U") ' выполняется поиск 
'   Call POISK_ZAD("O3", 1, "07", "08", "09", "12", "OSK_обновл") ' добавил 27.11.2020 поиск 
'   Call POISK_ZAD("N1",17, "FA", "IM", "OT", "DR", "стат до 2019")  ' и вывод в 4 таблицы
'   Call POISK_ZAD("PR",17, "FA", "IM", "OT", "DR", "стат 2019")  ' добавил преступления 2019
'   Call POISK_ZAD("N2", 1, "07", "08", "09", "12", "лица до 2019")
'   Call POISK_ZAD("RZ", 1, "3" , "4" , "5" , "8" , "розыск настоящий")
'   Call POISK_ZAD("ZZ", 4, "FAM","IMJ",""  , "FAM","база пробная мульти")     
   Call POISK_ZAD_V("SP", 1, "F1", "F2", "F3", "DR", "поселенцы")     

   Response.Write "<HR><p><font size='3' > "  

'   if KOL_ZAP=0 then
'   Response.Write "<A href='add_IF.asp?F1=" & F1 & "&F2=" & F2 & "&F3=" & F3  & "&DR=" & DR& "'>введём новую запись ?</A>"
'   end if
'   Response.Write "<Table border=1 cellspacing=0 >"
'   Response.Write "<TD>&nbsp;" & F1 
'   Response.Write "<TD>&nbsp;" & F2
'   Response.Write "<TD>&nbsp;" & F3
'   Response.Write "<TD>&nbsp;&nbsp;" & DR
'   Response.Write "</TABLE>"

end if

if T="2" then
   if F3="" And F2="" then
      Response.Write "без № уд искать не буду !<BR>"
      Break
   end if
'   Call POISK_ZAD("O2", 4, "O", "G", "N", "32",  "OSK_U") ' выполняется поиск 
'   Call POISK_ZAD("O3", 4, "O", "G", "N", "32",  "OSK_обновле") ' выполняется 'поиск 
'   Call POISK_ZAD("N1", 1, "O", "G", "N", "PNK", "стат до 2019") ' и вывод в 4 таблицы
'   Call POISK_ZAD("PR", 1, "O", "G", "N", "PNK", "стат 2019") ' 
'   Call POISK_ZAD("N2", 1, "O", "G", "N", "05",  "лица до 2019")
'   Call POISK_ZAD("Rz", 1, "O", "G", "N", "2",   "розыск настоящий")
'   Call POISK_ZAD("AI", 1, "O", "G", "N", "04",  "Архив")
   Call POISK_ZAD("SP", 1, "5", "NLD", "AN", "7",  "Архив")
end if

db.Close
Set db = Nothing
'Response.Write "<HR><p><font size='1' > "  
Response.Write "время поиска "  
'response.write(DateDiff("h",time,tim_start) & " час ")
'response.write(DateDiff("n",time,tim_start) & " минут ")
response.write(DateDiff("s",time,tim_start) & " секунд <br />")
Response.Write " </font> "  


%>
<A href="sp--form.asp?T=1"><font color=Red>Выход</FONT></A><BR>
<!--<FONT size=1 Color=Tan>просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.06.2020 </FONT> -->
</HTML>
                                              