<!DOCTYPE html>
<html>
<HEAD>
<TITLE>проба Закрытие окна </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<script type="text/javascript">
ID=window.setTimeout("Update();",2000);
function Update(){
   window.close();
   }
</script>
<!--#INCLUDE FILE="vvo-vid.inc"       ' для корректировки и добавления фото -->
<script src = "sweetalert.min.js"></script>
<link rel="stylesheet" href="style.css">

</HEAD>
<BODY>
</body>
</html>

<%

sub VID_SP
   Response.Write "<Table><TR><TD vAlign=Top width=350>"
   Response.Write "<Table border=1 bgcolor=White>"
   Call Vid_Reks ("область",       "OBL", "SP_OBL") 
   Call Vid_Rek  ("фамилия",       "F1") 
   Call Vid_Rek  ("имя",           "F2") 
   Call Vid_Rek  ("отчество",      "F3") 
   Call Vid_RekD ("год рожд",      "DR") 
   Call Vid_Rek  ("№ личного дела","NLD")
   Call Vid_RekD ("дата прибытия", "13")
   Call Vid_RekD ("дата снятия",   "DS")
   Call Vid_Rek  ("архивный №",    "AN")
   Call Vid_Reks ("пол",           "5", "SP_POL")
   Call Vid_Reks ("нац",           "6", "SP_NAC") 
   Call Vid_Reks ("категория",     "7", "SP_KAT") 
   Call Vid_Rek  ("место рождения","8")
   Call Vid_Rek  ("адрес",         "10")

   Call Vid_Reks ("секретность",   "9", "SP_sek") 
   Call Vid_Rek  ("примечание",    "11")
   Call Vid_Rek  ("IP",            "IP")
   Call Vid_RekD ("дата записи АТОС",   "ZZZ")
   Call Vid_RekD ("дата корректир АТОС","KKK")
   Call Vid_RekB ("дата записи2",   "AT_V_D")
   Call Vid_RekB ("дата корректир2","AT_K_D")

   Response.Write "</TABLE>"
   NP=N-1 'алгоритм просмотра предыдущей записи начало--------------
   if NP=0 then  NP=NP+1 
   Response.Write chr(13) & "<BR><A href='vvo-vid-" & Name_Zad & ".asp?N=" & NP & "'><DIV align=left>предыдущая запись</DIV></A><BR>" 

   NK=N+1 'алгоритм просмотра следующей записи начало--------------
   if NK=0 then  NK=NK+1
   Response.Write chr(13) & "<A href='vvo-vid-" & Name_Zad & ".asp?N=" & NK & "'><DIV align=left>следующая запись</DIV></A><BR>" 
'   Response.Write chr(13) & "<A href='edit-" & Name_Zad & ".asp?N=" & N & "'><DIV align=left> Корректировка</DIV></A><BR>" 
   Response.Write "<TD valign=top>"

   Response.Write "</TABLE>"
end sub


sub VID_F
   ID=rs.Fields("AT_0").value
   ZAPROS = "Select * From ATOC_SP_F where "
   ZAPROS = ZAPROS & "AT_1='" & ID & "';"
   On Error Resume Next          ' включает контроль ошибок
   Set RS_F = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка запроса в таблице F<BR>" & ZAPROS & "<BR>"  
      exit sub
   end if
                                         
 if RS_F.eof then
    Response.Write "<A href='foto_add.asp?ID=" & ID & "'> добавить фото?</A><BR>"
 else   
      Response.Write "<Table border=1 cellspacing=0>"
      Do while NOT RS_F.EOF 
         N_F     = RS_F.Fields("AT_0").value 
         F_NAME  = RS_F.Fields("F_NAME").value
         FIL_SIZE  = RS_F.Fields("FIL_SIZE").value
         PATH ="/foto_gala/"
         FIL_NAME = PATH&F_NAME     ' эТИ ДВЕ СТРОЧКИ ПОМЕНЯЙ И SMOTR.ASP 
         Response.Write "<TR><TD><EMBED src='" & FIL_NAME&"'  height=510 width =820></EMBED>" 
         RS_F.MoveNext          
         Response.Write "</TABLE>"
      Loop                    
         Response.Write "<p><A href='del_f.asp?ID=" &ID & "'><DIV align=center> удалить фото</DIV></A></p>" 'надо закоммертировать для всех кроме 
 end if
   Response.Write "</DIV>"
   Set RS_F = Nothing
end sub

call OPEN_BAZA("SP","")
ID = Request.QueryString("AT_1")
if NOT Rs.EOF then
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=0 cellpadding=5>"
    Response.Write "<TD valign=top>"
    call VID_SP
    Response.Write "<TD valign=top>"   

    call VID_F
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
else

   Response.Write "SP не прочиталось !"

end if

call CLOSE_BAZA

'N1 Уголовная статистика
' Довлетбаев З.С.                     ' 2019.02.20      
' 2020.09.16 работа с архивом

%>
