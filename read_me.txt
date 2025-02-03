
вызов в HTML

<A href="program.htm">текст</A><BR>
<A href="program.asp">текст</A><BR>
<A href="program.asp?N=1">текст</A><BR>
<A href="program.asp?N=365&PARAM=BOBRPA">текст</A><BR>
<A href="program.asp?N=365&PARAM=BOBRPA&T=2">текст</A><BR>


вызов в ASP


Response.Write "<A href="program.htm">текст</A><BR>"
Response.Write "<A href="program.asp">текст</A><BR>"
N=365
Response.Write "<A href="program.asp?N=" & N & ">текст</A><BR>"

N=444
TXT="текст"
Response.Write "<A href="program.asp?N=" & N & ">" & TXT & "</A><BR>"

N=444
PARM=ИВАН
TXT="текст"
Response.Write "<A href="program.asp?N=" & N & "&PARM=" & PARM & ">" & TXT & "</A><BR>"



