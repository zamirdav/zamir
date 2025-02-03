
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>save_SP</title>
</HEAD>
<BODY>
<!--#INCLUDE FILE="save.inc"-->
<%
' записывает в базу
' Довлетбаев З.С.
' 2019.02.05

OPEN_BAZA("SP")

if NOT Rs.EOF then
   Response.Write "<Table border=1>"
   Response.Write "<TR><TD>поле<TD>в базе<TD>новое значение<TD>запрос"
   call DATA_KORR(0)
   Call write_Rek ("область",           "OBL") 
   Call write_Rek ("фамилия",   	"F1")
   Call write_Rek ("имя",       	"F2")
   Call write_Rek ("отчество",  	"F3")
   Call write_Rek ("дата рожд", 	"DR")
   Call write_Rek ("пол", 		"5")
   Call write_Rek ("национальность", 	"6")
   Call write_Rek ("категория", 	"7")
   Call write_Rek ("место рождения", 	"8")   
   Call write_Rek ("секретность",       "9")
   Call write_Rek ("адрес",       	"10")
   Call write_Rek ("примечание",       	"11")
   Call write_Rek ("дата прибытия", 	"13")
   Call write_Rek ("арх.номер", 	"AN")  
   Call write_Rek ("N личного дела", 	"NLD")
   Call write_Rek ("дата снятия с учета","DS")
   Call write_Rek ("IP",          	"IP")
   Call write_Rek ("дата записи", 	"ZZZ")
   Call write_Rek ("дата корректировки","KKK")


   call DATA_KORR(1)


   Response.Write "</Table>"
   Set rs = Nothing
end if

CLOSE_BAZA

%>
