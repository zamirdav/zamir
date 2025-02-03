<%

dim DB

function OPEN_DSN                    ' функция открытия базы данных
   On Error Resume Next
'   DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\TEMP\WWW\lico-a.mdb"
'  DSN = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\TEMP\WWW\lico-a.mdb"
   DSN ="ATOC1"
   Set DB = Server.CreateObject("ADODB.Connection") 
       DB.Open DSN   
   if Err.Number<>0 then
      Response.Write "Не открылся DSN=" & DSN & "<BR>"
      Err.Clear
      OPEN_DSN=1
      exit function
   end if
   OPEN_DSN=0
end function


sub LICO_ADD(F1)           ' добавление записи в базу данных
    On Error Resume Next
    ZAPROS = "INSERT INTO LICO (F1, DR, Adr) values ('"&F1&"','19630428',1);"
    Response.Write ZAPROS & "<BR>"
    DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
end sub




function REK_SLV(s_NAM,KOD)      ' поиск и выборка по словарю
   Set RS_RS = DB.Execute("select * from " & s_NAM & " where Id=" & KOD & ";")
   if RS_RS.EOF then
      A=""
   else
      A=RS_RS.Fields("Txt").value
   end if
   Set RS_RS = Nothing
   REK_SLV=A 
end function







sub LICO_VID_ODIN(ID)        ' поиск и вывод в таблицу из в баз данных
    On Error Resume Next
    ZAPROS = "select * from LICO where L_ID=" & ID & ";"
'   Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
'   Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2>"
    if NOT RS.EOF then
       ID=RS.Fields("L_ID").value 
       Response.Write "<TR><TD>Фамилия <TD>" & RS.Fields("F1").value 
       Response.Write "<TR><TD>Имя     <TD>" & RS.Fields("F2").value 
       Response.Write "<TR><TD>Отчество<TD>" & RS.Fields("F3").value 
       Response.Write "<TR><TD>DR      <TD>" & RS.Fields("DR").value 
       Response.Write "<TR><TD>пол     <TD>" & REK_SLV("s_PL",RS.Fields("PL").value)  
       Response.Write "<TR><TD>м/р     <TD>" & RS.Fields("MR").value 
       Response.Write "<TR><TD>ку      <TD>" & REK_SLV("s_KU",RS.Fields("KU").value)  
       Response.Write "<TR><TD>адр     <TD>" & RS.Fields("Adr").value 
       Response.Write "<TR><TD>дк      <TD>" & RS.Fields("K_DAT").value 
       Response.Write "<TR><TD>IP      <TD>" & RS.Fields("IP").value 
    end if                    
    Set RS = Nothing
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
end sub




function SET_FIO(Z)  
   SET_FIO = uCase(mid(Z,1,1)) & LCase(mid(Z,2,31))
end function                               




function VID_SLV(s_NAM,GOM)       ' скролинг словаря
   Response.Write chr(10)  
   Set RS_S = DB.Execute("select * from " & s_NAM & " where Txt<>'-';")
   A=""
   Do while NOT RS_S.EOF 
      Kod=RS_S.Fields("ID").value
      Zna=RS_S.Fields("Txt").value
      T=""
      if GOM&"-" = Kod&"-" then 
         T="selected"
'        IP_set=rs.Fields("IP").value
      end if
      A=A & "<option " & T & " value='" & KOD & "'>" & Zna & "</option>" & chr(10)  
      RS_S.MoveNext          
   Loop                    
   Set RS_S = Nothing
   VID_SLV=A & chr(10) ' & "</SELECT>" 
end function





sub FOTO_SAVE(L_ID,TIP,PAT,NAM,OPIS)     '
    On Error Resume Next  'On Error используется в неструктурированной обработке ошибок
    if Err.Number<>0 then
       Response.Write "FOTO_SAVE на входе Err.Number="& Err.Number & "<BR>"
       Err.Clear
       exit sub
    end if

    Set objStream = Server.CreateObject("ADODB.Stream") 'чтения, записи и управления потоком двоич данных или текста.
    objStream.Type = 1
    objStream.Open
    objStream.LoadFromFile(PAT & NAM) 'перезаписывает текущее содержимое объекта Stream данными, считанными из файла

    Set RS = CreateObject("ADODB.Recordset")
    RS.Open "FOTO", DB, 3, 3               ' FOTO хранится тута
    RS.AddNew
    RS.Fields("LIC").Value      = L_ID
    RS.Fields("TIP").Value      = TIP
    RS.Fields("Fil_Name").Value = OPIS
    RS.Fields("Fil_Size").Value = objStream.Size
    RS.Fields("TIP_CONT").Value = "image/jpeg"
    RS.Fields("FOT").Value = objStream.Read
    RS.Update
    RS.Close
    Set RS = Nothing
    Response.Write "----- фото " & L_ID & " записано " & TIP & " " & NAM & " " & OPIS & "<BR>" & chr(10)
    objStream.Close
    Set objStream = Nothing
end sub



%>
