
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>edit SP </title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->
<%                                                                                                         	
' ���������� �.�.
' 2023.09.14
'�� �������� ������ ��������� ������� �� ���: 0550 102614 ���������� ����� ������������ �(�)
call OPEN_BAZA("SP")
if NOT Rs.EOF then
   Response.Write "<FORM action='save-"&Name_Zad&".asp'>"
   Response.Write "<input type='hidden' name='ADRES' value=" & N & ">"
   Response.Write "<input type='hidden' name='T' value=" & TipBaza & ">"
   Response.Write "<CENTER>"
   Response.Write "<Table><TR><TD vAlign=Top width=350>"
   Response.Write "<Table border=1 bgcolor=lavender>"    '  cellspacing=0
   Call edit_RekS("�������",            "OBL", "SP_OBL") 
   Call edit_Rek ("�������",   		"F1")
   Call edit_Rek ("���",       		"F2")
   Call edit_Rek ("��������",  		"F3")
   Call edit_Rek ("���� ����", 		"DR")
   Call edit_Rek ("N ������� ����", 	"NLD")
   Call edit_Rek("���� ��������", 	"13")
   Call edit_Rek("���� ������ � �����",	"DS")
   Call edit_Rek ("���.�����", 		"AN")  
   Call edit_RekS("���", 		"5", "N2_POL")
   Call edit_RekS("��������������", 	"6", "SP_NAC")
   Call edit_RekS("���������", 		"7", "SP_KAT")
   Call edit_Rek ("����� ��������", 	"8")   
   Call edit_Rek ("�����",       	"10")

   Call edit_RekS("�����������",         "9", "SP_SEK")
   Call edit_Rek ("����������",       	"11")
   Call edit_Rek ("IP",          	"IP")
   Call edit_Rek("���� ������", 	"ZZZ")
   Call edit_Rek("���� �������������",	"KKK")
   Response.Write "</TABLE>"
   Response.Write "<TD vAlign=Top>"
   Response.Write "</TABLE>"
   Response.Write "</TABLE><BR>"
'   Response.Write "<DIV align=right> ���������� <input type='checkbox' name='VID_SAVE' value='V'></DIV>"
   Response.Write "<input type='SUBMIT' value='��������'>"
   Response.Write "</CENTER>"
   Response.Write "</FORM>"
'ID = Request.QueryString("N")
else
   Response.Write "SP �� ����������� !"
end if

call CLOSE_BAZA

%>
