
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>save_SP</title>
</HEAD>
<BODY>
<!--#INCLUDE FILE="save.inc"-->
<%
' ���������� � ����
' ���������� �.�.
' 2019.02.05

OPEN_BAZA("SP")

if NOT Rs.EOF then
   Response.Write "<Table border=1>"
   Response.Write "<TR><TD>����<TD>� ����<TD>����� ��������<TD>������"
   call DATA_KORR(0)
   Call write_Rek ("�������",           "OBL") 
   Call write_Rek ("�������",   	"F1")
   Call write_Rek ("���",       	"F2")
   Call write_Rek ("��������",  	"F3")
   Call write_Rek ("���� ����", 	"DR")
   Call write_Rek ("���", 		"5")
   Call write_Rek ("��������������", 	"6")
   Call write_Rek ("���������", 	"7")
   Call write_Rek ("����� ��������", 	"8")   
   Call write_Rek ("�����������",       "9")
   Call write_Rek ("�����",       	"10")
   Call write_Rek ("����������",       	"11")
   Call write_Rek ("���� ��������", 	"13")
   Call write_Rek ("���.�����", 	"AN")  
   Call write_Rek ("N ������� ����", 	"NLD")
   Call write_Rek ("���� ������ � �����","DS")
   Call write_Rek ("IP",          	"IP")
   Call write_Rek ("���� ������", 	"ZZZ")
   Call write_Rek ("���� �������������","KKK")


   call DATA_KORR(1)


   Response.Write "</Table>"
   Set rs = Nothing
end if

CLOSE_BAZA

%>
