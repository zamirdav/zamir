<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>pro-stat0.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
2023  год по месяцам и по дате корректировки
<BR>
<%
'Довлетбаев З.С
'07.07.2021 сделал статистику
'выводит на экран по строкам органы а по столбцам годы по выбранной таблице
'Response.Buffer = False 
'Response.Flush = false
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "atoc1"   
Name_Zad="ATOC_SP_f"   


function Slv_Get(pSlv,znaRek) '---------- возвращает из словаря код и расшифровку 
   A=znaRek 
   On Error Resume Next
   Set rsSLV = db.Execute("Select * From slv_" & pSlv & " where KD=" & znaRek & " ;")
   if db.Errors.Count > 0 then
      A=A & " - не открылся"
   else
      if NOT rsSLV.EOF then A=A & " <FONT size=-1 color=DarkGreen>" & rsSLV.Fields("TX").value & "</FONT>"
      Set rsSLV = Nothing
   end if
   Slv_Get=A  ' возвращает из словаря код и расшифровку 
end function


dim MAC(999,999) ' обьявляем массив 

sub STAT_1    ' заполняет массив кол-ва МЕСЯЦЕВ по ДНЯМ--------------------------
   ZAPROS = "Select * From "&Name_Zad & " where AT_K_D >='20230101';"
'   Response.Write ZAPROS & "<BR>"
   On Error Resume Next          ' включает контроль ошибок
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка исполнения запроса<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      KOL=0
      Do while NOT Rs.EOF 
         KOL = KOL + 1 
         Dat=rs.Fields("AT_K_D").value 
         M = mid(Dat,5,2) 
         D = mid(Dat,7,2) 
'         Response.Write M & "__" & D & "<BR>"
         MAC(M,D)=MAC(M,D)+1 ' заполняет массив кол-ва органов по дням месяца
         Rs.MoveNext          
      Loop                    
end sub


sub PRINT(Z) ' выводит на экран таблицу----------------
    if Z=0 then Z=""
    Z="<A href='all-poisk.asp?T=2&P1=" & D & "&P2=" & M & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
    Response.Write "<TD>" & Z
end sub

   M=1              ' обнуляем массив---------------
   Do while M<13
      D=1
      Do while D<32
         MAC(M,D)=0
         D=D+1
      Loop                    
      M=M+1
   Loop                    



   call STAT_1
' считаем массив


' выводим таблицу ДНЕЙ по столбцам----------------
   Response.Write "<TABLE border=1><TR><TD>месяц\день"
   DK=1
   Do while DK<32                ' выводим номера колонок таблицы
      Response.Write "<TD>" & DK 
      DK=DK+1
   Loop                    
   Response.Write "<TD>" & "всего" 
M=1
ST=0
Do while M<13
Response.Write "<TR><TD>"& M
   D=1
   KL=0
   Do while D<32
      Response.Write "<TD>" & MAC(M,D) 
      KL=KL+MAC(M,D)
      D=D+1
   Loop
   Response.Write "<TD>" & KL
   M=M+1
Loop                    

D=1
SUM=0
Response.Write "<TR><TD>" & "Итого"
Do while D<32
   M=1
   ST=0
   Do while M<13
      ST=ST+MAC(M,D)
      M=M+1
   Loop
   Response.Write "<TD>" & ST 
   SUM=SUM+ST 
   D=D+1
Loop                    
Response.Write "<TD>" & SUM 

'   Response.Write "<TR>" & KL 
   
Response.Write "</TABLE>"
Response.Write "<BR>" 


' Response.Buffer = True 
%>
<BR>                  
<BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="назад.gif" border=0 alt="Выход"></A><BR>
<FONT size=1 Color=Tan>  просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.08.2020 </FONT><BR>
</BODY>
</HTML>

