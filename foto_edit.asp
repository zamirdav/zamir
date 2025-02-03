<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>foto_edit</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="initF.inc"-->
<%
'Response.Write "foto_edit<BR>"
'Response.Write "DavZS<BR>"
'Response.Write "2022.11.03<BR>"

ID = Request.QueryString("ID")



sub FOTO_DEL
   ZAPR_T = "Delete From FOTO where F_Id=" & ID & ";" 
   Response.Write "ZAPR_T " & ZAPR_T & "<BR>"
   DB.Execute(ZAPR_T)
end sub






sub FOTO_EDIT
    On Error Resume Next
    ZAPROS = "select * from FOTO where F_ID=" & ID & ";"
    Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=0 cellpadding=5><TR><TD>"
    Response.Write "<FORM action='lico_save.asp'>"
    Response.Write "<INPUT type='hidden' name='TABL' value='FOTO'>"
    Response.Write "<INPUT type='hidden' name='ID' value='" & ID & "'>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2>"
    if NOT RS.EOF then
       Response.Write "<TR><TD>тип     <TD><SELECT name='TIP' style='width: 90px'>" & VID_SLV("s_F_TIP",RS.Fields("TIP").value) 
       Response.Write "<TR><TD>прим    <TD><INPUT type='TEXT' name='FIL_NAME' size=100 value='" & RS.Fields("FIL_Name").value & "'>"
       Response.Write "<TR><TD>size    <TD>"  & RS.Fields("FIL_Size").value 
       Response.Write "<TR><TD>F_PRIM  <TD>"  & RS.Fields("F_Prim").value 
       Response.Write "<TR><TD>TIP_CONT<TD>"  & RS.Fields("TIP_CONT").value 
    end if                    
    Set RS = Nothing
    Response.Write "</TABLE>"
    Response.Write "<input type='SUBMIT' value='Сохранить'>"
    Response.Write "</FORM>"
    Response.Write "<TD>"
    call FOTO_VID(ID)
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
'   Response.Write "<A href='lico_edit.asp?ID=" & ID & "'>edit</A><BR>"
end sub




if OPEN_DSN=0 then
   TIP = Request.QueryString("T")
   if ID&"~" = "~" then 
      if TIP&"~" = "~" then 
         Response.Write "нету параметра.<BR>"
      else
'        call LICO_SAVE
      end if
   else
      if TIP = "FOTO_DEL"  then call FOTO_DEL
      if TIP = "FOTO_EDIT" then call FOTO_EDIT
   end if
end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
