<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>stat_sp_0.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="ATOC_SP_1"   
pSlv_BOK="sp_nac"   '������� ������� �������� ������ ���������� 
pSlv_SAP="sp_POL"   '������� ������� �������� ������ ���������� 
BOK="r6"       '������� ���� �������� ������ ����������
SAP="r5"



dim MAC(200,3) ' ��� ����� ����� ����������� �� ���-�� ����� �������� �������� 
'STROK=200 '���������� ����� �������
'KOLONOK=3 '���������� ����� �������





function Slv_Get(pSlv,znaRek)             ' ���������� �� ������� ��� � ����������� 
   A=znaRek 
   On Error Resume Next
   Set rsSLV = db.Execute("Select * From slv_" & pSlv & " where KD=" & znaRek & ";")
   if db.Errors.Count > 0 then
      A=A & " - �� ��������"
   else
      if NOT rsSLV.EOF then A=A & " <FONT size=-1 color=DarkGreen>" & rsSLV.Fields("TX").value & "</FONT>"
      Set rsSLV = Nothing
   end if
   Slv_Get=A
end function




sub STAT_1
   ZAPROS = "Select * From "&Name_Zad&" order by "&BOK&","&SAP&";"
'      Response.Write ZAPROS & "<BR>"
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
         S=rs.Fields(BOK).value 
         K=rs.Fields(SAP).value 
         if S<1 then S=200    	'STROK
         if K<1 then K=3      	'KOLONOK
         if S>200 then S=200 	'STROK
         if K>3 then K=3	'KOLONOK
         MAC(S,K)=MAC(S,K)+1
         Rs.MoveNext          
      Loop                    
end sub


sub PRINT(Z)
    if Z=0 then Z=""
    Response.Write "<TD>" & Z
end sub


   S=1
   Do while S<200    ' 	STROK
      K=1
      Do while K<3   '	KOLONOK
         MAC(S,K)=0
         K=K+1
      Loop                    
      S=S+1
   Loop                    

   call STAT_1
   Response.Write "���������� ���������� �� �������������� � ���� "& "<BR><BR>"
   Response.Write "<TABLE border=><TR><TD>��������������\���"
   K=1
   Do while K<3		'KOLONOK
      Response.Write "<TD>" & Slv_Get(pSlv_SAP,K)
 
      K=K+1
   Loop                    
   S=1
   Do while S<4		'STROK S<200
      DA=0                                    
      K=1
      Do while K<3		'KOLONOK
         if MAC(S,K)>0 then DA=1
         K=K+1
      Loop                    
      if DA>0 then
         Response.Write "<TR><TD>" & Slv_Get(pSlv_BOK,S)
         K=1
         Do while K<3			'KOLONOK
            call PRINT(MAC(S,K))
            K=K+1
         Loop                    
     end if
     S=S+1
   Loop                    
   Response.Write "</TABLE>"

%>
<BR>                  
<BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="�����.gif" border=0 alt="�����"></A><BR>
<FONT size=1 Color=Tan>  �������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.08.2020 </FONT><BR>
</BODY>
</HTML>

