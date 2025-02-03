<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   
KOL_ZAP=0
' попытка работы с сессиями 15.07.2021
   KOT_VES = Session("VES")   
'  Response.Write "KOT_VES=" & KOT_VES & "<BR>"
   if KOT_VES&"~" = "~" then
      IP_ADR= Request.ServerVariables("REMOTE_ADDR")     
'     Response.Write "ADDR=" & IP_ADR & "<BR>"
      ZAPR_P="select * from IP_SPISOK where IP='" & IP_ADR & "';"
      Set RSP = DB.Execute(ZAPR_P)  
      if Err.Number<>0 then
         Response.Write "Запрос " & ZAPR_P & " не выполнился.<BR>"
         Response.Write "DSN=" & DSN & "<A href='PAROL/'>.</A><BR>"
         Err.Clear
      else
         KOT_FIO="вход"
         KOT_VES=0
         Do while NOT RSP.EOF 
            Session("VES")=rsP.Fields("KOT").value
            Session("NAM")=rsP.Fields("FIO").value
            KOT_VES = Session("VES")
'           Response.Write "KOT_VES=" & KOT_VES & "<BR>"
            RSP.MoveNext          
         Loop 
         Set RSP = Nothing
      end if                                              
   end if
   KOT_FIO = Session("NAM")   



sub Stat_IP_Add(Zad,Vid,Str)                            ' ------- 2020.06.03 --------------------------
   IP  = Request("REMOTE_ADDR")
   Dat = Date()
   Dat = mid(Dat,7,4) & "." & mid(Dat,4,3) & mid(Dat,1,2) 
   Tim = Time()
   if Len(Tim) < 8 then Tim="0" & Tim
   Par = Request.QueryString("T")
   Str1= Replace(Str,"'","''")
   ZAPR="insert into Stat_IP (Dat,Tim,Zad,Vid,IP,Par,Str) values ('" &Dat&"','"&Tim&"','"&Zad&"','"&Vid&"','"&IP&"','"&Par&"','"&Str1&"');"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "ошибка вставки записи в таблицу <BR>"   
      Response.Write "ошибка"&db.Errors.Count&"<BR>"
      exit sub
   end if
   Set rStat = Nothing
end sub


F1 = Request.QueryString("P1")
F2 = Request.QueryString("P2")
F3 = Request.QueryString("P3")
DR = Request.QueryString("P4")

V1 = MID(F1,1,4)
V2 = MID(F2,1,2)
V3 = MID(F3,1,2)
V4 = MID(DR,1,2)


sub POISK_ZAD(nZad,nFile,nF1,nF2,nF3,nDR,nTEXT) 
   Response.Write "<CENTER>"
   VO=" like "
   S=mid(F1,1,1)
   if (S>="0")and(S<="9") then VO="="
   ZAPROS = "Select * From ATOC_"&nZad&"_"&nFile&" where r"&nF1& VO & "'" & F1 & "%'"
   if F2<>"" then ZAPROS = ZAPROS &                " and r"&nF2& VO & "'" & F2 & "%'"
   if F3<>"" then ZAPROS = ZAPROS &                " and r"&nF3& VO & "'" & F3 & "%'"
   if DR<>"" then ZAPROS = ZAPROS &                " and r"&nDR& VO & "'" & DR & "%'"
   ZAPROS = ZAPROS & " order by r"&nDR&";"      
'   Response.Write ZAPROS & "<BR>"
   call Stat_IP_Add(nZad,"IskUD",ZAPROS)            ' запись в таблицу Stat_IP
   On Error Resume Next          ' включает контроль ошибок
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка открытия таблицы 1<BR>"   
      Response.Write ZAPROS & "<BR>"
      Response.Write "</CENTER>"
      exit sub
   end if
   if Rs.eof then
      Response.Write "ничего нету.<BR>"
      call Stat_IP_Add(nZad,"IskUD",ZAPROS)            ' запись в таблицу Stat_IP
'      Response.Write ZAPROS & "<BR>"
   else
	Response.Write "<BR>" & nZad & "  просмотр " & "<FONT size=+3>"

      Response.Write "<Table border=1 cellspacing=0 >"
      KOL_ZAP=0
      Do while NOT Rs.EOF 
         KOL_ZAP = KOL_ZAP + 1 
         N=rs.Fields("AT_0").value 
         Response.Write "<TR><TD align=right><A href='vid-"&nZad&".asp?N=" & N & "&F=" & nFile & "'>" & KOL_ZAP & "&nbsp;</A>"
         Response.Write "<TD>&nbsp;" & rs.Fields("r"&nF1).value 
         Response.Write "<TD>&nbsp;" & rs.Fields("r"&nF2).value 
         Response.Write "<TD>&nbsp;" & rs.Fields("r"&nF3).value 
         Response.Write "<TD>&nbsp;&nbsp;" & rs.Fields("r"&nDR).value 
         Rs.MoveNext          
      Loop                    
      Response.Write "</TABLE>"
   end if
   Set Rs = Nothing
   Response.Write "</CENTER>"
end sub

sub POISK_ZAD_V(nZad,nFile,nF1,nF2,nF3,nDR,nTEXT) 
   Response.Write "<CENTER>"
   VO=" like "
   S=mid(V1,1,1)
   if (S>="0")and(S<="9") then VO="="
   ZAPROS = "Select * From ATOC_"&nZad&"_"&nFile&" where r"&nF1& VO & "'" & V1 & "%'"
   if V2<>"" then ZAPROS = ZAPROS &                " and r"&nF2& VO & "'" & V2 & "%'"
   if V3<>"" then ZAPROS = ZAPROS &                " and r"&nF3& VO & "'" & V3 & "%'"
   if DR<>"" then ZAPROS = ZAPROS &                " and r"&nDR& VO & "'" & V4 & "%'"
   ZAPROS = ZAPROS & " order by r"&nDR&";"      
'   Response.Write ZAPROS & "<BR>"
   call Stat_IP_Add(nZad,"IskUD",ZAPROS)            ' запись в таблицу Stat_IP
   On Error Resume Next          ' включает контроль ошибок
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка открытия таблицы 1<BR>"   
      Response.Write ZAPROS & "<BR>"
      Response.Write "</CENTER>"
      exit sub
   end if
   if Rs.eof then
      Response.Write "ничего нету.<BR>"
      call Stat_IP_Add(nZad,"IskUD",ZAPROS)            ' запись в таблицу Stat_IP
'      Response.Write ZAPROS & "<BR>"
   else
'	Response.Write "<BR>" & nZad & "  похожие записи уже есть " & "<FONT size=+3>"

      Response.Write "<Table border=1 cellspacing=0 >"
      KOL_ZAP=0
      Do while NOT Rs.EOF 
         KOL_ZAP = KOL_ZAP + 1 
         N=rs.Fields("AT_0").value 
         Response.Write "<TR><TD align=right><A href='vid-"&nZad&".asp?N=" & N & "&F=" & nFile & "'>" & KOL_ZAP & "&nbsp;</A>"
         Response.Write "<TD>&nbsp;" & rs.Fields("r"&nF1).value 
         Response.Write "<TD>&nbsp;" & rs.Fields("r"&nF2).value 
         Response.Write "<TD>&nbsp;" & rs.Fields("r"&nF3).value 
         Response.Write "<TD>&nbsp;&nbsp;" & rs.Fields("r"&nDR).value 
         Rs.MoveNext          
      Loop                    
      Response.Write "</TABLE>"
   end if
   Set Rs = Nothing
   Response.Write "</CENTER>"
end sub




%>
