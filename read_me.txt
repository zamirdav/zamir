
����� � HTML

<A href="program.htm">�����</A><BR>
<A href="program.asp">�����</A><BR>
<A href="program.asp?N=1">�����</A><BR>
<A href="program.asp?N=365&PARAM=BOBRPA">�����</A><BR>
<A href="program.asp?N=365&PARAM=BOBRPA&T=2">�����</A><BR>


����� � ASP


Response.Write "<A href="program.htm">�����</A><BR>"
Response.Write "<A href="program.asp">�����</A><BR>"
N=365
Response.Write "<A href="program.asp?N=" & N & ">�����</A><BR>"

N=444
TXT="�����"
Response.Write "<A href="program.asp?N=" & N & ">" & TXT & "</A><BR>"

N=444
PARM=����
TXT="�����"
Response.Write "<A href="program.asp?N=" & N & "&PARM=" & PARM & ">" & TXT & "</A><BR>"



