<HTML>
<HEAD>
   <META http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>ATOC.asp</TITLE>
</HEAD>
<BODY BGCOLOR="lightblue">
<BR>
<BR>
<% 
N = Request.QueryString("N")    
if N = "1" then 
   Response.Redirect "vvo-form.asp?T=1"
end if 
if N = "2" then 
   Response.Redirect "vvo-form.asp?T=2"
end if 
if N = "0" then 
   Response.Write "�������� ������.<BR>"
   P = Request.QueryString("P")    
   D=Date()
   D1=mid(D,1,2)
   if P=D1 then
      Response.Write "��������� !<BR>"
      Response.Write "<A href='vvo-form.asp?T=1'> �� ��� </A><BR>"
      Response.Write "<A href='vvo-form.asp?T=2'> �� �� </A><BR>"
      Response.Write "<BR>"
      Response.Write "<BR>"
      Response.Write "<BR>"
      Response.Write "<A href='PAROL/'> parol </A><BR>"
   else
      Response.Write "�� ������ !<BR>"
'      Response.Write D1 & "<BR>"
   end if
end if 
%>
<BR>
<BR>
<BR>
<BR>
<FONT SIZE=1 color=lightblue>27.01.2019 ���������� �.�.</FONT><BR>
<FONT SIZE=1 color=lightblue>29.06.2020 ���������� �.�.</FONT><BR>
<A href="index.htm"><IMG src="�����.gif" border=0 alt="�����" width="64" height="32"></A><BR>
<BODY> 
<HTML>
