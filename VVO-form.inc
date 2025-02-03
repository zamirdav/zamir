<%
' 2022.11.09
'07_09_2023  начал программу ОСК для ввода информации из регионов
' использовал наработки из выселенцев
'14_09_2023 пришлось для ввода информациии и перехода его в меню ввода создать из vid.inc vvo_vid.inc

'
DIM T
sub FORMA_FIO
%>
<!--
<h6>программу создал <BR>
Довлетбаев Замир Самарбекович<BR>
в 2022 году <BR>
-->
<CENTER>
<CENTER>
<table BORDER=5>
   <td WIDTH="270" height="180" BGCOLOR="lightblue" valign=top>
   <CENTER>
   <FONT size=5> Вводите  </FONT>
   <FORM action="VVO-poisk.asp">        
      <input type="HIDDEN" name="T" value="1">
      <TABLE >
         <TR>
         <TD>
	 Фамилия    <TD><input type="TEXT" name="P1" value="">
         <TR><TD>Имя        <TD><input type="TEXT" name="P2" value="">
         <TR><TD>Отчество   <TD><input type="TEXT" name="P3" value="">
         <TR><TD>Дата рожд. <TD><input type="TEXT" name="P4" value="1">
         <TR><TD> </TD><TD><BR>
         <input type="SUBMIT" value="Провека на двойников"></TD></TR>
         <TD>
         <A href="../index.htm" ><align=right>Выход</A></TD></TR>
      </TABLE>
   </FORM>
   </CENTER>
</table>
</CENTER>
<BR>
<BR>
<%
end sub      


sub FORMA_UD
%>
поиск по № УД все базы<BR>
<BR>
<BR>
<CENTER>
<CENTER>
<table BORDER=5>
   <tr>
   <td WIDTH="340" height="140" BGCOLOR="lightblue" valign=top>
   <CENTER><br>
   <FONT size=5> какой архивный номер ? </FONT><BR>
   <BR>
   <FORM action="VVO-poisk.asp">
      <input type="HIDDEN" name="T" value="2">
      <TABLE>
	   <TR><TD>архивный №   <TD><input type="TEXT" name="P3" value="50">
	   <TR><TD>личный №     <TD><input type="TEXT" name="P2" value="">
	   <TR><TD> </TD><TD><BR>
           <input type="SUBMIT" value="Искать"></TD></TR>
           <TR><TD> </TD><TD><BR>
           <A href="../index.htm">Выход</A></TD></TR>

      </TABLE>
   </FORM>
   </CENTER>
</table>
</CENTER>
<BR>
<BR>
<%
end sub

%>
