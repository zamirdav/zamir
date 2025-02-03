<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>stat_SP_1.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   
' генератор отчетов- для месяцев и дней
' генератор отчетов- создает множество одиночных тадиц с одинаковыми бок и шапок 
' которые отличаются лишь доп поиском допустим "where r5=1 и обьед в табл SELECT или LEFT JOIN  
sub sp_count(bok,shap,file,file_out,poisk)  ' готовит одну таблицу для общего отчета из бок
' и count(шапка обычно AT_0) (будет во всех других табл)+ поиск  
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       'если есть табл с таким именем удаляем
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка 1 в DROP <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TABLE "& file_out&"(bok int,shap int);"  'создаем табл
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка 2 CREATE TABLE <BR>"   
      exit sub
   end if                                   
'ниже вставляем записи в табл из селекта с добавлением poisk допустим "where r5=1" 
   ZAPR = "insert into "& file_out&"(bok,shap) SELECT MONTH("&bok&"),count("&shap&") from "&file&" "&poisk&" Group by YEAR("&bok&") Order by YEAR("&bok&");"
'   Response.Write "ZAPR=" & ZAPR & "<BR>"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "ошибка вставки записи в таблицу3 <BR>"   
      exit sub
   end if 
   Set rStat = Nothing
end sub

sub sp_console(file_out)  ' вывод на экран
   ZAPR = "Select * From "&file_out&" GROUP BY bok ORDER BY bok ;" 
'  Response.Write ZAPR & "<BR>"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №=9" & db.Errors.Count & "<BR>"
      Response.Write "ZAPROS=" & ZAPR & "<BR>"                  
      Response.Write "ошибка вставки записи в таблицу3 <BR>"   
      exit sub
   end if 
   if rStat.eof then
      Response.Write "ничего нету.<BR>"
      Response.Write ZAPR & "<BR>"
   else
'     ВВЕДИ НАЗВАНИЕ ТАБЛИЦЫ
      Response.Write "месяцы 2023/ дни<BR>"                  
'     Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "<Table border=1 cellspacing=0>"
         Response.Write "<TR><TD> год "
'         Response.Write "<TD> shap0 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 1 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 2 "  
         Response.Write "<TD> 3 "  
         Response.Write "<TD> 4 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 5 "  
         Response.Write "<TD> 6 "  
         Response.Write "<TD> 7 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 8 "  
         Response.Write "<TD> 9 "  
         Response.Write "<TD> 10 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 11 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 12 "  
         Response.Write "<TD> 13 "  
         Response.Write "<TD> 14 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 15 "  
         Response.Write "<TD> 16 "  
         Response.Write "<TD> 17 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 18 "  
         Response.Write "<TD> 19 "  
         Response.Write "<TD> 20 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 21 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 22 "  
         Response.Write "<TD> 23 "  
         Response.Write "<TD> 24 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 25 "  
         Response.Write "<TD> 26 "  
         Response.Write "<TD> 27 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 28 "  
         Response.Write "<TD> 29 "  
         Response.Write "<TD> 30 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> 31 "  

      Do while NOT rStat.EOF      
         Response.Write "<TR><TD> "& rStat.Fields("bok").value
'         Response.Write "<TD> " & rStat.Fields("shap0").value 
         Response.Write "<TD> " & rStat.Fields("shap1").value 
         Response.Write "<TD> " & rStat.Fields("shap2").value 
         Response.Write "<TD> " & rStat.Fields("shap3").value 
         Response.Write "<TD> " & rStat.Fields("shap4").value 
         Response.Write "<TD> " & rStat.Fields("shap5").value 
         Response.Write "<TD> " & rStat.Fields("shap6").value       
         Response.Write "<TD> " & rStat.Fields("shap7").value 
         Response.Write "<TD> " & rStat.Fields("shap8").value       
         Response.Write "<TD> " & rStat.Fields("shap9").value 
         Response.Write "<TD> " & rStat.Fields("shap10").value
         Response.Write "<TD> " & rStat.Fields("shap11").value 
         Response.Write "<TD> " & rStat.Fields("shap12").value 
         Response.Write "<TD> " & rStat.Fields("shap13").value 
         Response.Write "<TD> " & rStat.Fields("shap14").value 
         Response.Write "<TD> " & rStat.Fields("shap15").value 
         Response.Write "<TD> " & rStat.Fields("shap16").value       
         Response.Write "<TD> " & rStat.Fields("shap17").value 
         Response.Write "<TD> " & rStat.Fields("shap18").value       
         Response.Write "<TD> " & rStat.Fields("shap19").value 
         Response.Write "<TD> " & rStat.Fields("shap20").value
         Response.Write "<TD> " & rStat.Fields("shap22").value 
         Response.Write "<TD> " & rStat.Fields("shap23").value 
         Response.Write "<TD> " & rStat.Fields("shap24").value 
         Response.Write "<TD> " & rStat.Fields("shap25").value 
         Response.Write "<TD> " & rStat.Fields("shap26").value       
         Response.Write "<TD> " & rStat.Fields("shap27").value 
         Response.Write "<TD> " & rStat.Fields("shap28").value       
         Response.Write "<TD> " & rStat.Fields("shap29").value 
         Response.Write "<TD> " & rStat.Fields("shap30").value 
         Response.Write "<TD> " & rStat.Fields("shap31").value 
         rStat.MoveNext          
      Loop                    
      Response.Write "</TABLE>"
   end if
   Set rStat = Nothing
end sub

sub sp_union(file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file13,file14,file15,file16,file17,file18,file19,file20,file21,file22,file23,file24,file25,file26,file27,file28,file29,file30,file31,file_out) 
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       'если есть табл с таким именем удаляем
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №11=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка вставки записи в таблицу7 <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TABLE "& file_out&"(bok int,shap1 int,shap2 int,shap3 int,shap4 int,shap5 int,shap6 int,shap7 int,shap8 int,shap9 int,shap10 int,shap11 int,shap12 int)"
'   ZAPR =ZAPR+"shap8 int,shap9 int,shap10 int,shap11 int,shap12 int);"  'создаем табл
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №12=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка 2 CREATE TABLE <BR>"   
      exit sub
   end if    
    ZAPR = "insert into "&file_out&"(bok)"
    ZAPR =ZAPR+" SELECT "&file1&".bok FROM "&file1&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file2&".bok FROM "&file2&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file3&".bok FROM "&file3&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file4&".bok FROM "&file4&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file5&".bok FROM "&file5&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file6&".bok FROM "&file6&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file7&".bok FROM "&file7&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file8&".bok FROM "&file8&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file9&".bok FROM "&file9&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file10&".bok FROM "&file10&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file11&".bok FROM "&file11&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file12&".bok FROM "&file12&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file13&".bok FROM "&file13&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file14&".bok FROM "&file14&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file15&".bok FROM "&file15&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file16&".bok FROM "&file16&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file17&".bok FROM "&file17&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file18&".bok FROM "&file18&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file19&".bok FROM "&file19&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file20&".bok FROM "&file20&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file21&".bok FROM "&file21&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file22&".bok FROM "&file22&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file23&".bok FROM "&file23&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file24&".bok FROM "&file24&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file25&".bok FROM "&file25&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file26&".bok FROM "&file26&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file27&".bok FROM "&file27&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file28&".bok FROM "&file28&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file29&".bok FROM "&file29&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file30&".bok FROM "&file30&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file31&".bok FROM "&file31&" GROUP BY "&file31&".bok ORDER BY "&file31&".bok;"

'   Response.Write "ZAPR=" & ZAPR & "<BR>"
'    Response.Write "ZAPR=" "&file3&".bok & "<BR>"

   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №13 =" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка вставки записи в таблицу  <BR>"   
      exit sub
   end if  
   Set rStat = Nothing
end sub

sub sp_count_plus(file0,file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file13,file14,file15,file16,file17,file18,file19,file20,file21,file22,file23,file24,file25,file26,file27,file28,file29,file30,file31,file_out)  ' боковины шапки отчетов с услвием
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       'если есть табл с таким именем удаляем
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №11=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка вставки записи в таблицу7 <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TABLE "& file_out&"(bok int,shap1 int,shap2 int,shap3 int,shap4 int,shap5 int,shap6 int,shap7 int,shap8 int,shap9 int,shap10 int,shap11 int,shap12 int,shap13 int,shap14 int,shap15 int,shap16 int,shap17 int,shap18 int,shap19 int,shap20 int,shap21 int,shap22 int,shap23 int,shap24 int,shap25 int,shap26 int,shap27 int,shap28 int,shap29 int,shap30 int,shap31 int)"
'   ZAPR =ZAPR+"shap8 int,shap9 int,shap10 int,shap11 int,shap12 int);"  'создаем табл
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №12=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"                                                    
      Response.Write "ошибка CREATE TABLE <BR>"   
      exit sub
   end if    
    ZAPR = "insert into "&file_out&"(bok,shap1,shap2,shap3,shap4,shap5,shap6,shap7,shap8,shap9,shap10,shap11,shap12,shap13,shap14,shap15,shap16,shap17,shap18,shap19,shap20,shap21,shap22,shap23,shap24,shap25,shap26,shap27,shap28,shap29,shap30,shap31)"
    ZAPR =ZAPR+" SELECT "&file0&".bok,"&file1&".shap,"&file2&".shap,"&file3&".shap,"&file4&".shap"
    ZAPR =ZAPR+","&file5&".shap,"&file6&".shap,"&file7&".shap,"&file8&".shap,"&file9&".shap,"&file10&".shap"
    ZAPR =ZAPR+","&file11&".shap,"&file12&".shap,"&file13&".shap,"&file14&".shap,"&file15&".shap,"&file16&".shap"
    ZAPR =ZAPR+","&file17&".shap,"&file18&".shap,"&file19&".shap,"&file20&".shap,"&file21&".shap,"&file22&".shap"
    ZAPR =ZAPR+","&file23&".shap,"&file24&".shap,"&file25&".shap,"&file26&".shap,"&file27&".shap,"&file28&".shap"
    ZAPR =ZAPR+","&file29&".shap,"&file30&".shap,"&file31&".shap FROM "&file0&""
    ZAPR =ZAPR+" LEFT JOIN "&file1&" ON "&file1&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file2&" ON "&file2&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file3&" ON "&file3&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file4&" ON "&file4&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file5&" ON "&file5&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file6&" ON "&file6&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file7&" ON "&file7&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file8&" ON "&file8&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file9&" ON "&file9&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file10&" ON "&file10&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file11&" ON "&file11&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file12&" ON "&file12&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file13&" ON "&file13&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file14&" ON "&file14&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file15&" ON "&file15&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file16&" ON "&file16&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file17&" ON "&file17&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file18&" ON "&file18&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file19&" ON "&file19&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file20&" ON "&file20&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file21&" ON "&file21&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file22&" ON "&file22&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file23&" ON "&file23&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file24&" ON "&file24&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file25&" ON "&file25&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file26&" ON "&file26&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file27&" ON "&file27&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file28&" ON "&file28&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file29&" ON "&file29&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file30&" ON "&file30&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file31&" ON "&file31&".bok="&file0&".bok"

'    ZAPR =ZAPR+"  GROUP BY "&file0&".bok ORDER BY "&file0&".bok;"  
'  Response.Write "ZAPR=" & ZAPR & "<BR>"
'    Response.Write "ZAPR=" "&file3&".bok & "<BR>"

   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ошибка №14 =" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка вставки записи в таблицу  <BR>"   
      exit sub
   end if  
   Set rStat = Nothing
end sub
'подпрограмма просто соединяет один столбец всех таблиц для получ ощего для всех столбца и применения JOIN

p=1
do WHIlE P<32
   poisk=" where YEAR(AT_K_D)=2023 AND DAY(AT_K_D)="&P   ' модно вставлять поиск допустим "where r5=1"
   call sp_count("rzzz","AT_0","atoc_sp_1","D_"&p,poisk) 'табл из бок(будет во всех других табл) и count(шапка обычно AT_0)+ поиск  
   p=p+1
Loop
'call sp_console("D_1")                     
'call sp_console("D_2")                     
'call sp_console("D_3")                     
call sp_union("D_1","D_2","D_3","D_4","D_5","D_6","D_7","D_8","D_9","D_10","D_11","D_12","D_13","D_14","D_15","D_16","D_17","D_18","D_19","D_20","D_21","D_22","D_23","D_24","D_25","D_26","D_27","D_28","D_29","D_30","D_31","D_pr")
call sp_count_plus("D_pr","D_1","D_2","D_3","D_4","D_5","D_6","D_7","D_8","D_9","D_10","D_11","D_12","D_13","D_14","D_15","D_16","D_17","D_18","D_19","D_20","D_21","D_22","D_23","D_24","D_25","D_26","D_27","D_28","D_29","D_30","D_31","D_ok")

call sp_console("D_ok")                     
'call sp_count_plus("D_7","D_8","D_9","D_10","D_11","D_12","D_pr1")
'call sp_console("D_pr")                     

%>
<BR>                                                                                                                                  <BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="назад.gif" border=0 alt="Выход"></A><BR>                                                           <FONT size=1 Color=Tan>  просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.08.2020 </FONT><BR>
</BODY>
</HTML>

