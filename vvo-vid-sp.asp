<!DOCTYPE html>
<html>
<HEAD>
<TITLE>����� �������� ���� </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<script type="text/javascript">
ID=window.setTimeout("Update();",2000);
function Update(){
   window.close();
   }
</script>
<!--#INCLUDE FILE="vvo-vid.inc"       ' ��� ������������� � ���������� ���� -->
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
   Call Vid_Reks ("�������",       "OBL", "SP_OBL") 
   Call Vid_Rek  ("�������",       "F1") 
   Call Vid_Rek  ("���",           "F2") 
   Call Vid_Rek  ("��������",      "F3") 
   Call Vid_RekD ("��� ����",      "DR") 
   Call Vid_Rek  ("� ������� ����","NLD")
   Call Vid_RekD ("���� ��������", "13")
   Call Vid_RekD ("���� ������",   "DS")
   Call Vid_Rek  ("�������� �",    "AN")
   Call Vid_Reks ("���",           "5", "SP_POL")
   Call Vid_Reks ("���",           "6", "SP_NAC") 
   Call Vid_Reks ("���������",     "7", "SP_KAT") 
   Call Vid_Rek  ("����� ��������","8")
   Call Vid_Rek  ("�����",         "10")

   Call Vid_Reks ("�����������",   "9", "SP_sek") 
   Call Vid_Rek  ("����������",    "11")
   Call Vid_Rek  ("IP",            "IP")
   Call Vid_RekD ("���� ������ ����",   "ZZZ")
   Call Vid_RekD ("���� ��������� ����","KKK")
   Call Vid_RekB ("���� ������2",   "AT_V_D")
   Call Vid_RekB ("���� ���������2","AT_K_D")

   Response.Write "</TABLE>"
   NP=N-1 '�������� ��������� ���������� ������ ������--------------
   if NP=0 then  NP=NP+1 
   Response.Write chr(13) & "<BR><A href='vvo-vid-" & Name_Zad & ".asp?N=" & NP & "'><DIV align=left>���������� ������</DIV></A><BR>" 

   NK=N+1 '�������� ��������� ��������� ������ ������--------------
   if NK=0 then  NK=NK+1
   Response.Write chr(13) & "<A href='vvo-vid-" & Name_Zad & ".asp?N=" & NK & "'><DIV align=left>��������� ������</DIV></A><BR>" 
'   Response.Write chr(13) & "<A href='edit-" & Name_Zad & ".asp?N=" & N & "'><DIV align=left> �������������</DIV></A><BR>" 
   Response.Write "<TD valign=top>"

   Response.Write "</TABLE>"
end sub


sub VID_F
   ID=rs.Fields("AT_0").value
   ZAPROS = "Select * From ATOC_SP_F where "
   ZAPROS = ZAPROS & "AT_1='" & ID & "';"
   On Error Resume Next          ' �������� �������� ������
   Set RS_F = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "������ ������� � ������� F<BR>" & ZAPROS & "<BR>"  
      exit sub
   end if
                                         
 if RS_F.eof then
    Response.Write "<A href='foto_add.asp?ID=" & ID & "'> �������� ����?</A><BR>"
 else   
      Response.Write "<Table border=1 cellspacing=0>"
      Do while NOT RS_F.EOF 
         N_F     = RS_F.Fields("AT_0").value 
         F_NAME  = RS_F.Fields("F_NAME").value
         FIL_SIZE  = RS_F.Fields("FIL_SIZE").value
         PATH ="/foto_gala/"
         FIL_NAME = PATH&F_NAME     ' ��� ��� ������� ������� � SMOTR.ASP 
         Response.Write "<TR><TD><EMBED src='" & FIL_NAME&"'  height=510 width =820></EMBED>" 
         RS_F.MoveNext          
         Response.Write "</TABLE>"
      Loop                    
         Response.Write "<p><A href='del_f.asp?ID=" &ID & "'><DIV align=center> ������� ����</DIV></A></p>" '���� ���������������� ��� ���� ����� 
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

   Response.Write "SP �� ����������� !"

end if

call CLOSE_BAZA

'N1 ��������� ����������
' ���������� �.�.                     ' 2019.02.20      
' 2020.09.16 ������ � �������

%>
