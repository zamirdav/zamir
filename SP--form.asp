<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>ATOC.asp zapros ALL</TITLE>
</HEAD>
<style>
        body { background: url(../../video/кыргызприрода8.jpg);
        background-repeat: no-repeat;
       	background-size: cover;

       }
</style>                                   	
<!-- прозрачный текст -->
<style>
.transparent {
    font-family: Tahoma, sans-serif;
    font-weight: bold;
    font-size: 30px;
    line-height: 50px;
    text-transform: uppercase;
    background: #FFF;
    color: #FFF;
    mix-blend-mode: multiply;
    padding: 10px 20px;
    display: inline-block;
    /* text-shadow: 0 0 8px rgba(0,0,0,5), 0 2px 4px rgba(0,0,0,0.7); */
}
</style>

<!--#INCLUDE FILE="sp--form.inc"-->
<%
T = Request.QueryString("T") 

if T="1" then
   call FORMA_FIO
end if
if T="2" then
   call FORMA_UD
end if

%>
<BR>
<div class="transparent">
<h6>
(C)© по вопросам работы программы звонить по тел: 0550 102614 <BR>
программу создал в 2022.11.09 Довлетбаев Замир Самарбекович


</h6>
</body>
</html>

