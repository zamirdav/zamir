<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>pro-stat0.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
2023  ��� �� ������� � �� ���� �������������
<BR>
<%
'���������� �.�
'07.07.2021 ������ ����������
'������� �� ����� �� ������� ������ � �� �������� ���� �� ��������� �������
'Response.Buffer = False 
'Response.Flush = false
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "atoc1"   
Name_Zad="ATOC_SP_f"   


function Slv_Get(pSlv,znaRek) '---------- ���������� �� ������� ��� � ����������� 
   A=znaRek 
   On Error Resume Next
   Set rsSLV = db.Execute("Select * From slv_" & pSlv & " where KD=" & znaRek & " ;")
   if db.Errors.Count > 0 then
      A=A & " - �� ��������"
   else
      if NOT rsSLV.EOF then A=A & " <FONT size=-1 color=DarkGreen>" & rsSLV.Fields("TX").value & "</FONT>"
      Set rsSLV = Nothing
   end if
   Slv_Get=A  ' ���������� �� ������� ��� � ����������� 
end function


dim MAC(999,999) ' ��������� ������ 

sub STAT_1    ' ��������� ������ ���-�� ������� �� ����--------------------------
   ZAPROS = "Select * From "&Name_Zad & " where AT_K_D >='20230101';"
'   Response.Write ZAPROS & "<BR>"
   On Error Resume Next          ' �������� �������� ������
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "������ ���������� �������<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      KOL=0
      Do while NOT Rs.EOF 
         KOL = KOL + 1 
         Dat=rs.Fields("AT_K_D").value 
         M = mid(Dat,5,2) 
         D = mid(Dat,7,2) 
'         Response.Write M & "__" & D & "<BR>"
         MAC(M,D)=MAC(M,D)+1 ' ��������� ������ ���-�� ������� �� ���� ������
         Rs.MoveNext          
      Loop                    
end sub


sub PRINT(Z) ' ������� �� ����� �������----------------
    if Z=0 then Z=""
    Z="<A href='all-poisk.asp?T=2&P1=" & D & "&P2=" & M & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
    Response.Write "<TD>" & Z
end sub

   M=1              ' �������� ������---------------
   Do while M<13
      D=1
      Do while D<32
         MAC(M,D)=0
         D=D+1
      Loop                    
      M=M+1
   Loop                    



   call STAT_1
' ������� ������


' ������� ������� ���� �� ��������----------------
   Response.Write "<TABLE border=1><TR><TD>�����\����"
   DK=1
   Do while DK<32                ' ������� ������ ������� �������
      Response.Write "<TD>" & DK 
      DK=DK+1
   Loop                    
   Response.Write "<TD>" & "�����" 
M=1
ST=0
Do while M<13
Response.Write "<TR><TD>"& M
   D=1
   KL=0
   Do while D<32
      Response.Write "<TD>" & MAC(M,D) 
      KL=KL+MAC(M,D)
      D=D+1
   Loop
   Response.Write "<TD>" & KL
   M=M+1
Loop                    

D=1
SUM=0
Response.Write "<TR><TD>" & "�����"
Do while D<32
   M=1
   ST=0
   Do while M<13
      ST=ST+MAC(M,D)
      M=M+1
   Loop
   Response.Write "<TD>" & ST 
   SUM=SUM+ST 
   D=D+1
Loop                    
Response.Write "<TD>" & SUM 

'   Response.Write "<TR>" & KL 
   
Response.Write "</TABLE>"
Response.Write "<BR>" 


' Response.Buffer = True 
%>
<BR>                  
<BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="�����.gif" border=0 alt="�����"></A><BR>
<FONT size=1 Color=Tan>  �������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.08.2020 </FONT><BR>
</BODY>
</HTML>

