<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>edit F</title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->

<% '��������� ���������� ���� INSERT � ��������� �������
'Response.Write "���������� ����� ������������ <BR>"<!--09112022�(�)-->
'ADRES = Request.QueryString("ADRES") 
' ������������ 19092023 ���� ������� ��� ��������� ����� ���� ��� ������� ATOC_AI_F �� ATOC_SP_F 
'�� � ������� ���� ������� �������������� ������� ����� 

sub RECORD_ADD
   ZAPROS="select * from ATOC_sp_F where AT_1='" & ID & "' ;"

   On Error Resume Next  
   Set RS = DB.Execute(ZAPROS) 
   if db.Errors.Count > 0 then
      Response.Write "������ ���������� ������� " & ZAPROS & "<BR>" 
      Set RS = Nothing
   else
      if not RS.eof then
         ADR=RS.Fields("AT_0").value
         Response.Write "����� ������ ��� ���� <BR>"
'         Response.Write "<A href='foto_add.asp?ID=" & ID & "'> �������� ����?</A><BR>"
'         Response.Write "<A href='edit-SP.asp?N=" & ADR & "'>�������</A><BR>"   
         Set RS = Nothing
      else
' ������� ���� ����� ������    12 06 2023
         DAT=Date()
         if len(DAT)<10 then DAT="0" & DAT
         DAT=mid(DAT,7,4) & mid(DAT,4,2) & mid(DAT,1,2) 
         TIM=Time()
         if len(TIM)<8 then TIM="0" & TIM
         TIM=mid(TIM,1,2) & mid(TIM,4,2) & mid(TIM,7,2) 

         IP  = Request("REMOTE_ADDR") ' ��������� ����� � ������ ���� ���� 

         call FOTO_FORMA(ID)
         Z_VVOD1 = "insert into ATOC_sp_F(AT_1,FIL_NAME,F_NAME,FIL_SIZE,AT_V_D,AT_V_V,AT_V_KIM)"  
         Z_VVOD2 = " values('" & ID & "','" & FIL_NAME & "','" & F_NAME & "','" & FIL_SIZE & "','" & DAT & "','" & TIM & "','" & IP & "');" 
         Response.Write Z_VVOD1&Z_VVOD2 & "<BR>"
         Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
         if db.Errors.Count > 0 then
            Response.Write "������ ����������<BR>" &  db.Errors.Count
            Set RZ_Z = Nothing
         else
            Set RZ_Z = Nothing
            Set RS = DB.Execute(ZAPROS) 
            if db.Errors.Count > 0 then
               Response.Write "������ ���������� ������� " & ZAPROS & "<BR>" 
               Set RS = Nothing
            else
               if RS.eof then
                  Response.Write "������-�� ����� !<BR>"   
               else
                  ADR=RS.Fields("AT_0").value
                  Set RS = Nothing
                  Response.Redirect "vvo-form.asp?T=1"          '????????????????
'                 Response.Write "<A href='index.htm'> ����� </A><BR>"                                                                              end if
               end if
            end if
         end if
      end if
   end if
end sub                  


ID = Request.QueryString("ADRES")
FIL_NAME = Request.QueryString("FIL_NAME")
FIL_SIZE = Request.QueryString("FIL_SIZE")
F_NAME   = Request.QueryString("F_NAME")

'Response.Write "ID="&ID&"<BR>"
'Response.Write "FIL_NAME="&FIL_NAME&"<BR>"
'Response.Write "FIL_SIZE="&FIL_SIZE&"<BR>"
'Response.Write "F_NAME="&F_NAME&"<BR>"


call RECORD_ADD


'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="vvo-form.asp?T=1"> ����� </A><BR> 
<!--<A href="../index.htm"> ����� </A><BR>--> 




<%




%>

