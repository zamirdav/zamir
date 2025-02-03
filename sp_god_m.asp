<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>stat_SP_1.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
'Response.Write "Довлетбаев Замир Самарбекович <BR>"<!--09112022©(с)-->
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   
' генератор отчетов- для дат и месяцев 
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
   ZAPR = "insert into "& file_out&"(bok,shap) SELECT YEAR("&bok&"),count("&shap&") from "&file&" "&poisk&" Group by YEAR("&bok&") Order by YEAR("&bok&");"
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
      Response.Write "годы / месяцы <BR>"                  
'     Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "<Table border=1 cellspacing=0>"
         Response.Write "<TR><TD> год "
'         Response.Write "<TD> shap0 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> месяц1 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> месяц2 "  
         Response.Write "<TD> месяц3 "  
         Response.Write "<TD> месяц4 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> месяц5 "  
         Response.Write "<TD> месяц6 "  
         Response.Write "<TD> месяц7 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> месяц8 "  
         Response.Write "<TD> месяц9 "  
         Response.Write "<TD> месяц10 "   'поменяй название шапки и полуишь красота
         Response.Write "<TD> месяц11 "  
         Response.Write "<TD> месяц12 "  

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
         rStat.MoveNext          
      Loop                    
      Response.Write "</TABLE>"
   end if
   Set rStat = Nothing
end sub

sub sp_union(file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file_out) '
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
    ZAPR = "insert into "&file_out&"(bok)"',shap3,shap4,shap5,shap6)"
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
    ZAPR =ZAPR+" SELECT "&file12&".bok FROM "&file12&";"
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

sub sp_count_plus(file0,file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file_out) ',file3,file4,file5,file6,file_out) ',file8,file9,file10,file11,file12,file_out)  ' боковины шапки отчетов с услвием
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
    ZAPR = "insert into "&file_out&"(bok,shap1,shap2,shap3,shap4,shap5,shap6,shap7,shap8,shap9,shap10,shap11,shap12)"
    ZAPR =ZAPR+" SELECT "&file0&".bok,"&file1&".shap,"&file2&".shap,"&file3&".shap,"&file4&".shap"
    ZAPR =ZAPR+","&file5&".shap,"&file6&".shap,"&file7&".shap,"&file8&".shap,"&file9&".shap,"&file10&".shap"
    ZAPR =ZAPR+","&file11&".shap,"&file12&".shap FROM "&file0&""
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

'    ZAPR =ZAPR+"  GROUP BY "&file0&".bok ORDER BY "&file0&".bok;"  
'  Response.Write "ZAPR=" & ZAPR & "<BR>"
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
'подпрограмма просто соединяет один столбец всех таблиц для получ ощего для всех столбца и применения JOIN



poisk=" where month(AT_V_D)=1 "   ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_1",poisk) 'табл из бок(будет во всех других табл) и count(шапка обычно AT_0)+ поиск  

poisk=" where month(AT_V_D)=2" ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_2",poisk)
                      
poisk=" where month(AT_V_D)=3 "  ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_3",poisk) 

poisk=" where month(AT_V_D)=4 "   ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_4",poisk) 'табл из бок(будет во всех других табл) и count(шапка обычно AT_0)+ поиск  

poisk=" where month(AT_V_D)=5" ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_5",poisk)
                      
poisk=" where month(AT_V_D)=6 "  ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_6",poisk) 

poisk=" where month(AT_V_D)=7 "   ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_7",poisk) 'табл из бок(будет во всех других табл) и count(шапка обычно AT_0)+ поиск  

poisk=" where month(AT_V_D)=8" ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_8",poisk)
                      
poisk=" where month(AT_V_D)=9 "  ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_9",poisk) 

poisk=" where month(AT_V_D)=10 "   ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_10",poisk) 'табл из бок(будет во всех других табл) и count(шапка обычно AT_0)+ поиск  

poisk=" where month(AT_V_D)=11 " ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_11",poisk)
                      
poisk=" where month(AT_V_D)=12 "  ' модно вставлять поиск допустим "where r5=1"
call sp_count("AT_V_D","AT_0","atoc_sp_1","GM_12",poisk) 

'call sp_console("GM_1")                     
'call sp_console("GM_2")                     
'call sp_console("GM_3")                     
call sp_union("GM_1","GM_2","GM_3","GM_4","GM_5","GM_6","GM_7","GM_8","GM_9","GM_10","GM_11","GM_12","GM_pr")
call sp_count_plus("GM_pr","GM_1","GM_2","GM_3","GM_4","GM_5","GM_6","GM_7","GM_8","GM_9","GM_10","GM_11","GM_12","GM_ok")

call sp_console("GM_ok")                     
'call sp_count_plus("GM_7","GM_8","GM_9","GM_10","GM_11","GM_12","GM_pr1")
'call sp_console("GM_pr")                     

%>
<BR>                                                                                                                                  <BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="назад.gif" border=0 alt="Выход"></A><BR>                                                           <FONT size=1 Color=Tan>  просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.08.2020 </FONT><BR>
</BODY>
</HTML>

