<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>stat_SP_1.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   
' ��������� �������- ������� ��������� ��������� ����� � ����������� ��� � ����� 
' ������� ���������� ���� ��� ������� �������� "where r5=1 � ����� � ���� SELECT ��� JOIN  
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
   ZAPR = "CREATE TABLE "& file_out&"(bok int,shap int);"  '������� ����
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 2 CREATE TABLE <BR>"   
      exit sub
   end if                                   
'���� ��������� ������ � ���� �� ������� � ����������� poisk �������� "where r5=1" 
   ZAPR = "insert into "& file_out&"(bok,shap) SELECT "&bok&",count("&shap&") from "&file&" "&poisk&" Group by "&bok&" Order by "&bok&";"
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
   ZAPR = "Select * From "&file_out&" GROUP BY bok ORDER BY bok ;" 
   Response.Write ZAPR & "<BR>"
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
      Response.Write "ZAPR=" & ZAPR & "<BR>"                  
      Response.Write "<Table border=1 cellspacing=0>"
         Response.Write "<TR><TD> bok "
         Response.Write "<TD> shap1 "   '������� �������� ����� � ������� �������
         Response.Write "<TD> shap2 "  
         Response.Write "<TD> shap3 "  
      Do while NOT rStat.EOF      
         Response.Write "<TR><TD> "& rStat.Fields("bok").value
         Response.Write "<TD> " & rStat.Fields("shap1").value 
         Response.Write "<TD> " & rStat.Fields("shap2").value 
         Response.Write "<TD> " & rStat.Fields("shap3").value 

        rStat.MoveNext          
      Loop                    
      Response.Write "</TABLE>"
   end if
   Set rStat = Nothing
end sub

sub sp_count_plus(file1,file2,file3,file_out)  ' �������� ����� ������� � �������
   ZAPR = "DROP TABLE IF EXISTS "&file_out&";"       '���� ���� ���� � ����� ������ �������
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � �������7 <BR>"   
      exit sub
   end if
   ZAPR = "CREATE TABLE "& file_out&"(bok int,shap1 int,shap2 int,shap3 int);"  '������� ����
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ �=" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ 2 CREATE TABLE <BR>"   
      exit sub
   end if    
    ZAPR = "insert into "&file_out&"(bok,shap1,shap2,shap3)" ' ��� ����������� �������
    ZAPR =ZAPR+" SELECT "&file1&".bok,"&file1&".shap,"&file2&".shap,"&file3&".shap" 
    ZAPR =ZAPR+" FROM "&file1&","&file2&","&file3&"" 
    ZAPR =ZAPR+" WHERE "&file1&".bok="&file2&".bok AND "&file2&".bok="&file3&".bok" 
    ZAPR =ZAPR+" GROUP BY "&file1&".bok ORDER BY "&file1&".bok;"  
'   Response.Write "ZAPR=" & ZAPR & "<BR>"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "������ � =" & db.Errors.Count & "<BR>"
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � ������� BR>"   
      exit sub
   end if  
   Set rStat = Nothing
end sub



poisk=""   ' ����� ��������� ����� �������� "where r5=1"
call sp_count("r6","AT_0","atoc_sp_1","sp_1",poisk) '���� �� ���(����� �� ���� ������ ����) � count(����� ������ AT_0)+ �����  

poisk="where r5=1" ' ����� ��������� ����� �������� "where r5=1"
call sp_count("r6","AT_0","atoc_sp_1","sp_2",poisk)
                      
poisk="where r5=2"  ' ����� ��������� ����� �������� "where r5=1"
call sp_count("r6","AT_0","atoc_sp_1","sp_3",poisk) 
'call sp_console("sp_1")                     
'call sp_console("sp_2")                     
'call sp_console("sp_3")                     
call sp_count_plus("sp_1","sp_2","sp_3","sp_pr")
call sp_console("sp_pr")                     

%>
<BR>                                                                                                                                  <BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="�����.gif" border=0 alt="�����"></A><BR>                                                           <FONT size=1 Color=Tan>  �������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.08.2020 </FONT><BR>
</BODY>
</HTML>

