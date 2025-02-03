<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>Lico Create Table</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="init.inc"-->
<%
'Session("Name_Zad")="ZZ"
'Response.Write Session("Name_Zad") & "<BR>"
'Response.Write "Lico Create Table<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.21<BR>"



sub LICO_DELETE
    On Error Resume Next
    ZAPROS = "DROP TABLE LICO;"
    Response.Write ZAPROS & "<BR>"
    DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
end sub



sub LICO_CREATE
    On Error Resume Next
    ZAPROS = "CREATE TABLE LICO ("
    ZAPROS = ZAPROS & "L_ID COUNTER PRIMARY KEY, "
    ZAPROS = ZAPROS & "F1    TEXT(15)  "
    ZAPROS = ZAPROS & "F2    TEXT(15), "
    ZAPROS = ZAPROS & "F3    TEXT(15), "
    ZAPROS = ZAPROS & "DR    TEXT(8)), "
    ZAPROS = ZAPROS & "PL    INTEGER), "
    ZAPROS = ZAPROS & "Adr   INTEGER), "
    ZAPROS = ZAPROS & "X     TEXT(8)), "
    ZAPROS = ZAPROS & "K_DAT TEXT(6),  "
    ZAPROS = ZAPROS & "IP    TEXT(15));"
    Response.Write ZAPROS & "<BR>"
    DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"

    ZAPROS = "CREATE TABLE FOTO ("
    ZAPROS = ZAPROS & "F_ID COUNTER PRIMARY KEY, " 
    ZAPROS = ZAPROS & "LIC      INTEGER, " 
    ZAPROS = ZAPROS & "TIP      INTEGER, " 
    ZAPROS = ZAPROS & "FIL_Name CHAR,    " 
    ZAPROS = ZAPROS & "FIL_SIZE INTEGER, " 
    ZAPROS = ZAPROS & "F_PRIM   TEXT(15), " 
    ZAPROS = ZAPROS & "TIP_CONT TEXT(15), " 
    ZAPROS = ZAPROS & "FOT      IMAGE);" 
    Response.Write ZAPROS & "<BR>"
    DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
end sub


sub LICO_VID
    On Error Resume Next
    ZAPROS = "select * from LICO;"
    Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 bordercolor=blue>"
    N=0
    Do while NOT RS.EOF 
       N = N + 1     
       Response.Write "<TR><TD>" & N
       Response.Write "<TD>" & RS.Fields("L_ID").value 
       Response.Write "<TD>" & RS.Fields("F1").value 
       Response.Write "<TD>" & RS.Fields("F2").value 
       Response.Write "<TD>" & RS.Fields("DR").value 
       Response.Write "<TD>" & RS.Fields("Adr").value 
       RS.MoveNext          
    Loop                    
    Set RS = Nothing
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
end sub


if O PEN_DSN=0 then
   call LICO_DELETE
   call LICO_CREATE
   call LICO_ADD('Бобровский')
   call LICO_VID
else

end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
