
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>edit SP </title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->
<%                                                                                                         	
' Довлетбаев З.С.
' 2023.09.14
'по вопросам работы программы звонить по тел: 0550 102614 Довлетбаев Замир Самарбекович ©(с)
call OPEN_BAZA("SP")
if NOT Rs.EOF then
   Response.Write "<FORM action='save-"&Name_Zad&".asp'>"
   Response.Write "<input type='hidden' name='ADRES' value=" & N & ">"
   Response.Write "<input type='hidden' name='T' value=" & TipBaza & ">"
   Response.Write "<CENTER>"
   Response.Write "<Table><TR><TD vAlign=Top width=350>"
   Response.Write "<Table border=1 bgcolor=lavender>"    '  cellspacing=0
   Call edit_RekS("область",            "OBL", "SP_OBL") 
   Call edit_Rek ("фамилия",   		"F1")
   Call edit_Rek ("имя",       		"F2")
   Call edit_Rek ("отчество",  		"F3")
   Call edit_Rek ("дата рожд", 		"DR")
   Call edit_Rek ("N личного дела", 	"NLD")
   Call edit_Rek("дата прибытия", 	"13")
   Call edit_Rek("дата снятия с учета",	"DS")
   Call edit_Rek ("арх.номер", 		"AN")  
   Call edit_RekS("пол", 		"5", "N2_POL")
   Call edit_RekS("национальность", 	"6", "SP_NAC")
   Call edit_RekS("категория", 		"7", "SP_KAT")
   Call edit_Rek ("место рождения", 	"8")   
   Call edit_Rek ("адрес",       	"10")

   Call edit_RekS("секретность",         "9", "SP_SEK")
   Call edit_Rek ("примечание",       	"11")
   Call edit_Rek ("IP",          	"IP")
   Call edit_Rek("дата записи", 	"ZZZ")
   Call edit_Rek("дата корректировки",	"KKK")
   Response.Write "</TABLE>"
   Response.Write "<TD vAlign=Top>"
   Response.Write "</TABLE>"
   Response.Write "</TABLE><BR>"
'   Response.Write "<DIV align=right> посмотреть <input type='checkbox' name='VID_SAVE' value='V'></DIV>"
   Response.Write "<input type='SUBMIT' value='Записать'>"
   Response.Write "</CENTER>"
   Response.Write "</FORM>"
'ID = Request.QueryString("N")
else
   Response.Write "SP не прочиталось !"
end if

call CLOSE_BAZA

%>
