
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
   Response.Write "<H2> �������������� 4 </H2>"  

sub INDEX_ZAD(Tab,Ima,Ika)
    ZAPR_Ind=" SHOW INDEX FROM "&Tab& " Where Key_name='"&Ima&"';"
'    Response.Write"�������=" & ZAPR_Ind & "<BR>"

    Set rs = db.Execute(ZAPR_Ind)                        
   if db.Errors.Count > 0 then
      Response.Write "������ �������� ������� 1<BR>"   
      Response.Write ZAPR_IND & "<BR>"
      Response.Write "</CENTER>"
      exit sub
   end if
   if Rs.eof then
      Response.Write "��� ������� ������� " &Ima& " � " &Tab& "������ ������ <BR>"
   else
'      Response.Write "��� ���� ������� ������� " &Ima& " � " &Tab& "<BR>"
      exit sub
   end if

   ZAPR_IND = "Create Index "&Ima& " ON "&Tab&"("&Ika&");"   ' ��� ������ ?????
'   Response.Wrie ZAPROS & "<BR>"

   Response.Write "�������=" & ZAPR_IND & "<BR>"
On Error Resume Next
Set rs = db.Execute(ZAPR_IND)  
if Err.Number<>0 then
   Response.Write "������ �������������� "&nTab&" <BR>"
   Response.Write  Err.Number & ": " & Err.Description & "<HR>"
   Err.Clear
end if
Set rs = Nothing

end sub
%>
