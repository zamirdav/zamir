<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>stat_SP_1.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
'Response.Write "���������� ����� ������������ <BR>"<!--09112022�(�)-->
Tim_start=Time
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   
' ���������-��������� �������
' ����� �������� �������� �.� ����� ������ ����(��� ������� � ����)
' ������� ��������� ��������� ����� � ����������� ���(1-12) � ������� �������(1-31)  
' ������� ����������� � ���� �������   
' �������� ��������������� ������ ����� �������� ������ ������  

sub sp_bok(bok,file_out)  ' ������� ���� ������� �� �����(� ������ ������ 12 �������) ��� ������ ������ �� ���
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       '���� ���� ���� � ����� ������ �������
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 1 � DROP <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TABLE "& file_out&"(bok int);"  '������� ����
'CREATE TEMPORARY TABLE test.temp_table_memory (x int) ENGINE=MEMORY;
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 2 CREATE TABLE <BR>"   
      exit sub
   end if                                   
'���� ��������� ������ � ���� �� VALUES ("&p&")"  p-������
'm=p+1  ' p-������+1 ������ �������� ������ ����
'Response.Write "M=" & m & "<BR>"
p=1  ' p-������
do WHIlE p<32
   ZAPR = "insert into "& file_out&"() VALUES ("&p&");" ' ��������� ���� �������� p                                  
'   Response.Write "ZAPR=" & ZAPR & "<BR>"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �3=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "������ ������� ������ � �������3 <BR>"   
      exit sub
   end if 
   p=p+1
Loop
   Set rStat = Nothing
end sub

sub sp_count(bok,shap,file,file_out,poisk)  ' ������� ���� ������� ��� ������ ������ �� ���
' � count(����� ������ AT_0) (����� �� ���� ������ ����)+ �����  
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       '���� ���� ���� � ����� ������ �������
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 1 � DROP <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TEMPORARY TABLE "& file_out&"(bok int,shap int) ENGINE=MEMORY;"  '������� ����
'CREATE TEMPORARY TABLE test.temp_table_memory (x int) ENGINE=MEMORY;
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 2 CREATE TABLE <BR>"   
      exit sub
   end if                                   
'���� ��������� ������ � ���� �� ������� � ����������� poisk �������� "where r5=1" 
   ZAPR = "insert into "& file_out&"(bok,shap) SELECT DAY("&bok&"),count("&shap&") from "&file&" "&poisk&" Group by "&bok&" Order by "&bok&";"
'   Response.Write "ZAPR=" & ZAPR & "<BR>"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "������ ������� ������ � �������3 <BR>"   
      exit sub
   end if 
   Set rStat = Nothing
end sub

sub sp_console(file_out)  ' ����� �� �����
   ZAPR = "Select * From "&file_out&" GROUP BY bok;" 
'  Response.Write ZAPR & "<BR>"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=9" & db.Errors.Count & "<BR>"
      Response.Write "ZAPROS=" & ZAPR & "<BR>"                  
      Response.Write "������ ������� ������ � �������3 <BR>"   
      exit sub
   end if 
   if rStat.eof then
      Response.Write "������ ����.<BR>"
      Response.Write ZAPR & "<BR>"
   else
'     ����� �������� �������
      Response.Write "������ 2023/ ���<BR>"                  
'     Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "<Table border=1 cellspacing=0>"
         Response.Write "<TR><TD> ������ "
         Response.Write "<TD> 1 "   '������� �������� ����� � ������� �������
         Response.Write "<TD> 2 "  
         Response.Write "<TD> 3 "  
         Response.Write "<TD> 4 "   '������� �������� ����� � ������� �������
         Response.Write "<TD> 5 "  
         Response.Write "<TD> 6 "  
         Response.Write "<TD> 7 "   '������� �������� ����� � ������� �������
         Response.Write "<TD> 8 "  
         Response.Write "<TD> 9 "  
         Response.Write "<TD> 10 "   '������� �������� ����� � ������� �������
         Response.Write "<TD> 11 "   '������� �������� ����� � ������� �������
         Response.Write "<TD> 12 "  
'         Response.Write "<TD> 13 "  
'         Response.Write "<TD> 14 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 15 "  
'         Response.Write "<TD> 16 "  
'         Response.Write "<TD> 17 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 18 "  
'         Response.Write "<TD> 19 "  
'         Response.Write "<TD> 20 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 21 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 22 "  
'         Response.Write "<TD> 23 "  
'         Response.Write "<TD> 24 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 25 "  
'         Response.Write "<TD> 26 "  
'         Response.Write "<TD> 27 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 28 "  
'         Response.Write "<TD> 29 "  
'         Response.Write "<TD> 30 "   '������� �������� ����� � ������� �������
'         Response.Write "<TD> 31 "  


      Do while NOT rStat.EOF      
         Response.Write "<TR><TD> "& rStat.Fields("bok").value
'         Response.Write "<TD> " & rStat.Fields("shap").value 
         Response.Write "<TD> " & rStat.Fields("shap1").value 
         Response.Write "<TD> " & rStat.Fields("shap2").value 
         Response.Write "<TD> " & rStat.Fields("shap3").value 
         Response.Write "<TD> " & rStat.Fields("shap4").value 
         Response.Write "<TD> " & rStat.Fields("shap5").value 
         Response.Write "<TD> " & rStat.Fields("shap6").value       
         Response.Write "<TD> " & rStat.Fields("shap7").value 
         Response.Write "<TD> " & rStat.Fields("shap8").value       
         Response.Write "<TD> " & rStat.Fields("shap9").value 
         Response.Write "<TD> " & rStat.Fields("shap10").value
         Response.Write "<TD> " & rStat.Fields("shap11").value 
         Response.Write "<TD> " & rStat.Fields("shap12").value 
 '        Response.Write "<TD> " & rStat.Fields("shap13").value 
 '        Response.Write "<TD> " & rStat.Fields("shap14").value 
 '        Response.Write "<TD> " & rStat.Fields("shap15").value 
 '        Response.Write "<TD> " & rStat.Fields("shap16").value       
 '        Response.Write "<TD> " & rStat.Fields("shap17").value 
 '        Response.Write "<TD> " & rStat.Fields("shap18").value       
 '        Response.Write "<TD> " & rStat.Fields("shap19").value 
 '        Response.Write "<TD> " & rStat.Fields("shap20").value
 '        Response.Write "<TD> " & rStat.Fields("shap21").value
 '        Response.Write "<TD> " & rStat.Fields("shap22").value 
 '        Response.Write "<TD> " & rStat.Fields("shap23").value 
 '        Response.Write "<TD> " & rStat.Fields("shap24").value 
 '        Response.Write "<TD> " & rStat.Fields("shap25").value 
 '        Response.Write "<TD> " & rStat.Fields("shap26").value       
 '        Response.Write "<TD> " & rStat.Fields("shap27").value 
 '        Response.Write "<TD> " & rStat.Fields("shap28").value       
 '        Response.Write "<TD> " & rStat.Fields("shap29").value 
 '        Response.Write "<TD> " & rStat.Fields("shap30").value 
 '        Response.Write "<TD> " & rStat.Fields("shap31").value 
         rStat.MoveNext          
      Loop                    
      Response.Write "</TABLE>"
   end if
   Set rStat = Nothing
end sub
'sp_union ���������� ������� �� ��������� � ��������� 
sub sp_union(file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file_out) 
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       '���� ���� ���� � ����� ������ �������
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �11=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������7 <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TEMPORARY TABLE "& file_out&"(bok int) ENGINE=MEMORY" ',shap1 int,shap2 int,shap3 int,shap4 int,shap5 int,shap6 int,shap7 int,shap8 int,shap9 int,shap10 int,shap11 int,shap12 int)"
'CREATE TEMPORARY TABLE test.temp_table_memory (x int) ENGINE=MEMORY;
'   ZAPR =ZAPR+"shap8 int,shap9 int,shap10 int,shap11 int,shap12 int);"  '������� ����
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �12=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 2 CREATE TABLE <BR>"   
      exit sub
   end if    
    ZAPR = "insert into "&file_out&"(bok)"
    ZAPR =ZAPR+" SELECT "&file1&".bok FROM "&file1&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file2&".bok FROM "&file2&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file3&".bok FROM "&file3&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file4&".bok FROM "&file4&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file5&".bok FROM "&file5&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file6&".bok FROM "&file6&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file7&".bok FROM "&file7&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file8&".bok FROM "&file8&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file9&".bok FROM "&file9&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file10&".bok FROM "&file10&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file11&".bok FROM "&file11&" UNION ALL"
    ZAPR =ZAPR+" SELECT "&file12&".bok FROM "&file12&";"
'    ZAPR =ZAPR+" SELECT "&file13&".bok FROM "&file13&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file14&".bok FROM "&file14&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file15&".bok FROM "&file15&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file16&".bok FROM "&file16&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file17&".bok FROM "&file17&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file18&".bok FROM "&file18&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file19&".bok FROM "&file19&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file20&".bok FROM "&file20&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file21&".bok FROM "&file21&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file22&".bok FROM "&file22&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file23&".bok FROM "&file23&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file24&".bok FROM "&file24&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file25&".bok FROM "&file25&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file26&".bok FROM "&file26&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file27&".bok FROM "&file27&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file28&".bok FROM "&file28&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file29&".bok FROM "&file29&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file30&".bok FROM "&file30&" UNION ALL"
'    ZAPR =ZAPR+" SELECT "&file31&".bok FROM "&file31&";"
'   Response.Write "ZAPR=" & ZAPR & "<BR>"
'    Response.Write "ZAPR=" "&file3&".bok & "<BR>"

   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �13 =" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������  <BR>"   
      exit sub
   end if  
   Set rStat = Nothing
end sub

sub sp_union_plus(file,file_out) 
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       '���� ���� ���� � ����� ������ �������
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �11=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������7 <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TEMPORARY TABLE "& file_out&"(bok int) ENGINE=MEMORY;"  '������� ����
'CREATE TEMPORARY TABLE test.temp_table_memory (x int) ENGINE=MEMORY;
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �12=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 2 CREATE TABLE <BR>"   
      exit sub
   end if    
    ZAPR = "insert into "&file_out&"(bok) Select bok From "&file&" GROUP BY bok ORDER BY bok;" 
'   Response.Write "ZAPR=" & ZAPR & "<BR>"
'    Response.Write "ZAPR=" "&file3&".bok & "<BR>"

   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �15 =" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������  <BR>"   
      exit sub
   end if  
   Set rStat = Nothing
end sub

' ������� �� ���� ������ ���� ���� � ������� LEFT JOIN ������������� �����
sub sp_count_plus(file0,file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file_out)  ' �������� ����� ������� � �������
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       '���� ���� ���� � ����� ������ �������
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �11=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������7 <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TABLE "& file_out&"(bok int,shap1 int,shap2 int,shap3 int,shap4 int,shap5 int,shap6 int,shap7 int,shap8 int,shap9 int,shap10 int,shap11 int,shap12 int) ENGINE MEMORY; "

'   ZAPR =ZAPR+"shap8 int,shap9 int,shap10 int,shap11 int,shap12 int);"  '������� ����
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �12=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"                                                    
      Response.Write "������ CREATE TABLE <BR>"   
      exit sub
   end if    
    ZAPR = "insert into "&file_out&"(bok,shap1,shap2,shap3,shap4,shap5,shap6,shap7,shap8,shap9,shap10,shap11,shap12)"
    ZAPR =ZAPR+" SELECT "&file0&".bok,"&file1&".shap,"&file2&".shap,"&file3&".shap,"&file4&".shap"
    ZAPR =ZAPR+","&file5&".shap,"&file6&".shap,"&file7&".shap,"&file8&".shap,"&file9&".shap,"&file10&".shap"
    ZAPR =ZAPR+","&file11&".shap,"&file12&".shap FROM "&file0&""
    ZAPR =ZAPR+" LEFT JOIN "&file1&" ON "&file1&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file2&" ON "&file2&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file3&" ON "&file3&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file4&" ON "&file4&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file5&" ON "&file5&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file6&" ON "&file6&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file7&" ON "&file7&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file8&" ON "&file8&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file9&" ON "&file9&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file10&" ON "&file10&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file11&" ON "&file11&".bok="&file0&".bok"
    ZAPR =ZAPR+" LEFT JOIN "&file12&" ON "&file12&".bok="&file0&".bok"

'    ZAPR =ZAPR+"  GROUP BY "&file0&".bok ORDER BY "&file0&".bok;"  
'  Response.Write "ZAPR=" & ZAPR & "<BR>"
'    Response.Write "ZAPR=" "&file3&".bok & "<BR>"

   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ sp_count_plus_�14 =" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������  <BR>"   
      exit sub
   end if  
   Set rStat = Nothing
end sub


'----- ������--- ��������� ������ �� ������� � ���� ������ �����  
'----- sp_bok(p,"d_bok")'������� ������� �������� ������� 1<13 

'������� �������� ������� ������ �� ��������� p( � ������ ������ �� 12 �������) ������� ����� �� ���� ������ ����  
call sp_bok(bok,"d_bok")'������� ������� �������� �� � �����

'��������� � ����� ������� ������� �������� 
p=1
do WHIlE P<13
   poisk=" where YEAR(AT_V_D)=2023 AND MONTH(AT_V_D)="&P   ' ����� ��������� ����� �������� "where r5=1"
 ' ������� ���� ������� ��� ������ ������ �� ��� � count(����� ������ AT_0) (����� �� ���� ������ ����)+ �����  
   call sp_count("AT_V_D","AT_0","atoc_sp_1","D_"&p,poisk) '���� �� ���(����� �� ���� ������ ����) � count(����� ������ AT_0)+ �����  
   p=p+1
Loop
'call sp_console("D_bok")  ' ������� ������� �� �������                                   
'call sp_console("D_2")    ' ������� ������� �� �������                                                    
'call sp_console("D_3")    ' ������� ������� �� �������                                                    
'call sp_console("D_3")   ' ������� ������� �� �������                                                     

'sp_union �����-� ���� �� ����-� � �����-� � ���� ������� (����-�� ����� ������� ��� � �� ����-�� �������� ����������� � ������������ ����-� � ����-�)
'call sp_union("D_1","D_2","D_3","D_4","D_5","D_6","D_7","D_8","D_9","D_10","D_11","D_12","D_13","D_14","D_15","D_16","D_17","D_18","D_19","D_20","D_21","D_22","D_23","D_24","D_25","D_26","D_27","D_28","D_29","D_30","D_31","D_pr")
'call sp_union_plus("D_pr","D_pl") '���������� � ��������� ������� �� ��������� � ��������� D_pr � D_pl 

' ������� �� ���� ������ ���� ���� � ������� LEFT JOIN ������������� �����
call sp_count_plus("D_bok","D_1","D_2","D_3","D_4","D_5","D_6","D_7","D_8","D_9","D_10","D_11","D_12","D_ok")
call sp_console("D_ok")  ' ������� ������� �� �������                                                    

Response.Write "<HR><p><font size='1' > "  
Response.Write "����� ������ "  
response.write(DateDiff("s",time,tim_start) & " ������ <br />")
Response.Write " </font> "  

%>
<BR>                                                                                                                                  <BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="�����.gif" border=0 alt="�����"></A><BR>                                                           <FONT size=1 Color=Tan>  �������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.08.2020 </FONT><BR>
</BODY>
</HTML>

