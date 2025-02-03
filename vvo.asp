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
   Response.Write "Проверим пароль.<BR>"
   P = Request.QueryString("P")    
   D=Date()
   D1=mid(D,1,2)
   if P=D1 then
      Response.Write "Правильно !<BR>"
      Response.Write "<A href='vvo-form.asp?T=1'> по ФИО </A><BR>"
      Response.Write "<A href='vvo-form.asp?T=2'> по УД </A><BR>"
      Response.Write "<BR>"
      Response.Write "<BR>"
      Response.Write "<BR>"
      Response.Write "<A href='PAROL/'> parol </A><BR>"
   else
      Response.Write "Не угадал !<BR>"
'      Response.Write D1 & "<BR>"
   end if
end if 
%>
<BR>
<BR>
<BR>
<BR>
<FONT SIZE=1 color=lightblue>27.01.2019 Довлетбаев З.С.</FONT><BR>
<FONT SIZE=1 color=lightblue>29.06.2020 Довлетбаев З.С.</FONT><BR>
<A href="index.htm"><IMG src="назад.gif" border=0 alt="Выход" width="64" height="32"></A><BR>
<BODY> 
<HTML>
