
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>����� ALL</title>
</HEAD>

<!--#INCLUDE FILE="VVO-poisk.inc"--> 
<%
'07_09_2023  ����� ��������� ��� ��� ����� ���������� �� ��������
' ����������� ��������� �� ����������
'14_09_2023 �������� ��� ����� ����������� � �������� ��� � ���� ����� ������� �� vid.inc vvo_vid.inc

' 2020.08.26 ���������� ��������� �������
'2021.08.12 Or F1="" �������� ��� ���������� ������ 40 ��� ����������� �������� 
'
'
Tim_start=Time
'Response.Write "���������� ����� ������������ <BR>"<!--�(�)-->
'Response.Write "2019.01.24<BR>"
'27.11.2020 ������� ����� �� ��� �3


'Response.Write "<DIV align=right><A href='pass.asp'>&nbsp;" & KOT_FIO & "&nbsp;</A></DIV>"  

T = Request.QueryString("T") 

if T="1" then
   if F1&"~" = "~" then
      Response.Write "�� ��������� �������, ����� ��������� ����� ����� ������ ������ '�����'(<-) <BR>"   
'      Response.Write "<A href='vvo-form.asp?T=1'>�� ��������� �������, ���������?</A>"  
   else
      if F2&"~" = "~" then
         Response.Write "�� ��������� ���, ����� ��������� ����� ����� ������ ������ '�����'(<-) <BR>"   
'      Response.Write "<A href='vvo-form.asp?T=1'>�� ��������� ���, ���������?</A>"  
      else
         if DR&"~" = "~" then
           Response.Write "��� ���� ��������, ����� ��������� ����� ����� ������ ������ '�����'(<-) <BR>"   
'            Response.Write "<A href='vvo-form.asp?T=1'>�� ��������� ���� ��������, ���������?</A>"  
'            Break
         else
           Call POISK_ZAD_V("SP", 1, "F1", "F2", "F3", "DR", "���������")     
           Response.Write "<HR><p><font size='5' > "  
'          if KOL_ZAP=0 then
	   Response.Write "<A href='add_iF.asp?F1=" & F1 & "&F2=" & F2 & "&F3=" & F3  & "&DR=" & DR& "'>��� ���� ����� ������?</A>"
'	   end if
	   Response.Write "<Table border=1 cellspacing=0 >"
	   Response.Write "<TD>&nbsp;" & F1 
	   Response.Write "<TD>&nbsp;" & F2
	   Response.Write "<TD>&nbsp;" & F3
	   Response.Write "<TD>&nbsp;&nbsp;" & DR
	   Response.Write "</TABLE>"
         end if
      end if
   end if
'end if

end if

if T="2" then
   if F3="" And F2="" then
      Response.Write "��� � �� ������ �� ���� !<BR>"
      Break
   end if
'   Call POISK_ZAD("O2", 4, "O", "G", "N", "32",  "OSK_U") ' ����������� ����� 
'   Call POISK_ZAD("O3", 4, "O", "G", "N", "32",  "OSK_�������") ' ����������� '����� 
'   Call POISK_ZAD("N1", 1, "O", "G", "N", "PNK", "���� �� 2019") ' � ����� � 4 �������
'   Call POISK_ZAD("PR", 1, "O", "G", "N", "PNK", "���� 2019") ' 
'   Call POISK_ZAD("N2", 1, "O", "G", "N", "05",  "���� �� 2019")
'   Call POISK_ZAD("Rz", 1, "O", "G", "N", "2",   "������ ���������")
'   Call POISK_ZAD("AI", 1, "O", "G", "N", "04",  "�����")
   Call POISK_ZAD("SP", 1, "5", "NLD", "AN", "7",  "�����")
end if

db.Close
Set db = Nothing
Response.Write "<HR><p><font size='1' > "  
Response.Write "����� ������ "  
'response.write(DateDiff("h",time,tim_start) & " ��� ")
'response.write(DateDiff("n",time,tim_start) & " ����� ")
response.write(DateDiff("s",time,tim_start) & " ������ <br />")
Response.Write " </font> "  


%>
<A href="VVO-form.asp?T=1"><font size =+1 >����� ��� ������?</FONT></A><BR>
<!--<FONT size=1 Color=Tan>�������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.06.2020 </FONT> -->
<!--�(�)-->
</HTML>
                                              