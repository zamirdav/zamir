<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>edit AI</title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->

<% '��������� ���������� INSERT � �������
' ���� ������ � SP--poisk.asp
F1   = Request.QueryString("F1") 
F2   = Request.QueryString("F2") 
F3   = Request.QueryString("F3") 
DR   = Request.QueryString("DR")  
'ADRES = Request.QueryString("ADRES") 

Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  


sub RECORD_ADD
   ZAPROS="select * from ATOC_SP_1 where rF1='" & F1 & "' and rF2='" & F2 & "' and rF3='" & F3& "' and rDR='" & DR & "' ;"

   On Error Resume Next  
   Set RS = DB.Execute(ZAPROS) 
   if db.Errors.Count > 0 then
      Response.Write "������ ���������� ������� " & ZAPROS & "<BR>" 
      Set RS = Nothing
   else
      if not RS.eof then
         ADR=RS.Fields("AT_0").value
         Response.Write "����� ������ ��� ���� <BR>"  
         Response.Write "<A href='edit-SP.asp?N=" & ADR & "'>�������</A><BR>"   
         Set RS = Nothing
      else
'         if DR<>"" then DR = 1  ' ������� 22_12_2021
' ������� ���� ����� ������    23 05 2023
         DAT=Date()
         if len(DAT)<10 then DAT="0" & DAT
         DAT=mid(DAT,7,4) & mid(DAT,4,2) & mid(DAT,1,2) 
         TIM=Time()
         if len(TIM)<8 then TIM="0" & TIM
         TIM=mid(TIM,1,2) & mid(TIM,4,2) & mid(TIM,7,2) 
  
         IP  = Request("REMOTE_ADDR") ' ��������� ����� � ������ ���� ���� 

         Z_VVOD1 = "insert into ATOC_sp_1(rF1,rF2,rF3,rDR,AT_V_D,AT_V_V,AT_V_KIM)"  
         Z_VVOD2 = " values('" & F1 & "','" & F2 & "','" & F3 & "','" & DR &  "','" & DAT&  "','" & TIM &  "','" & IP & "');" 
         Response.Write Z_VVOD1&Z_VVOD2 & "<BR>"
         Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
         if db.Errors.Count > 0 then
            Response.Write "������ ����������<BR>"   
            Set RZ_Z = Nothing
         else
            Set RZ_Z = Nothing
            Response.Write Z_VVOD1&Z_VVOD2 & "<BR>"
            Set RS = DB.Execute(ZAPROS) 
            if db.Errors.Count > 0 then
               Response.Write "������ ���������� ������� " & ZAPROS & "<BR>" 
               Set RS = Nothing
            else
               if RS.eof then
'                  Response.Write ZAPROS & "<BR>"
                  Response.Write "������-�� ����� !<BR>"   
               else
                  ADR=RS.Fields("AT_0").value
                  Set RS = Nothing
                  Response.Redirect "edit-SP.asp?N=" & ADR 
                  Response.Write "���������<BR>"   
               end if
            end if
         end if
      end if
   end if
end sub



if F1&"~" = "~" then
   Response.Write "�� ��������� �������<BR>"   
else
   if F2&"~" = "~" then
      Response.Write "�� ��������� ��� <BR>"   
   else
      if DR&"~" = "~" then
         Response.Write "�� ��������� ��� �������� <BR>"   
      else
         call RECORD_ADD
      end if
   end if
end if

    
%>
