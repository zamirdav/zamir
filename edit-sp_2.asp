
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>edit SP </title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->

<%                                                                                                         	
' ���������� �.�.
' 2022.10.04
       
sub EDIT_SP
   Response.Write "<FORM action='save-"&Name_Zad&".asp'>"
   Response.Write "<input type='hidden' name='ADRES' value=" & N & ">"
   Response.Write "<input type='hidden' name='T' value=" & TipBaza & ">"
   Response.Write "<CENTER>"
   Response.Write "<Table><TR><TD vAlign=Top width=350>"
   Response.Write "<Table border=1 bgcolor=lavender>"    '  cellspacing=0
   Call Vid_Rek  ("�������",   		"F1")
   Call Vid_Rek  ("���",       		"F2")
   Call edit_Rek ("��������",  		"F3")
   Call Vid_Rek  ("���� ����", 		"DR")
   Call edit_RekS("���", 		"5", "N2_POL")
   Call edit_RekS("��������������", 	"6", "SP_NAC")
   Call edit_RekS("���������", 		"7", "SP_CAT")
   Call edit_Rek ("����� ��������", 	"8")   
   Call edit_RekS("�����������",         "9", "SP_SEK")
   Call edit_Rek ("�����",       	"10")
   Call edit_Rek ("����������",       	"11")
   Call edit_Rek("���� ��������", 	"13")
   Call edit_Rek ("���.�����", 		"AN")  
   Call edit_Rek ("N ������� ����", 	"NLD")
   Call edit_Rek("���� ������ � �����",	"DS")
   Call edit_Rek ("IP",          	"IP")
   Call edit_Rek("���� ������", 	"ZZZ")
   Call edit_Rek("���� �������������",	"KKK")

   Response.Write "</TABLE>"
   Response.Write "<TD vAlign=Top>"
   Response.Write "</TABLE>"
   Response.Write "</TABLE>"
   Response.Write "</CENTER>"
'   Response.Write "<DIV align=right> ���������� <input type='checkbox' name='VID_SAVE' value='V'></DIV>"
   Response.Write "<input type='SUBMIT' value='��������'>"
   Response.Write "</FORM>"
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
'      Response.Write "<Table border=1 cellspacing=0>"
      Do while NOT RS_F.EOF 
         N_F     = RS_F.Fields("AT_0").value 
         F_NAME  = RS_F.Fields("F_NAME").value
         FIL_SIZE  = RS_F.Fields("FIL_SIZE").value
         PATH ="http://10.30.0.53/ZAMIR_2022-10-04/foto_gala/"
         FIL_NAME = PATH&F_NAME     ' ��� ��� ������� ������� � SMOTR.ASP 
         Response.Write "<EMBED src='" & FIL_NAME&"'  height=580 width =820></EMBED>" 
         RS_F.MoveNext          
'         Response.Write "</TABLE>"
      Loop                    
         Response.Write "<A href='del_f.asp?ID=" &ID & "'><DIV align=center> ������� ����</DIV></A>" '���� ���������������� ��� ���� ����� 
   end if
'   Response.Write "</DIV>"
   Set RS_F = Nothing
end sub

call OPEN_BAZA("SP")
ID = Request.QueryString("N")
if NOT Rs.EOF then
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=0 cellpadding=5>"
    Response.Write "<TD valign=top width=350>"
    call EDIT_SP
    Response.Write "<!--#INCLUDE FILE='vid.inc'-->"
    Response.Write "<TD valign=top>"
    call VID_F
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
else
   Response.Write "SP �� ����������� !"
end if

call CLOSE_BAZA

%>
