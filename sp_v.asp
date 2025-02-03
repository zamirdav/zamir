<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>stat_SP_1.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
'Response.Write "Довлетбаев Замир Самарбекович <BR>"<!--09112022©(с)-->
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   
%>
DROP TABLE IF EXISTS sp_1; /* копирует в новую таблицу sp_1 все что в select */
CREATE TABLE sp_1(bok int,shap int);
insert into  sp_1(bok,shap)
  SELECT r6, count(AT_0) FROM atoc_sp_1
  GROUP BY r6
  order by r6;
  
DROP TABLE IF EXISTS sp_2; /* копирует в новую таблицу sp_2 все что в select */
CREATE TABLE sp_2(bok int,shap int);
insert into  sp_2(bok,shap)
  SELECT r6, count(at_0) FROM atoc_sp_1
  where r5=1
  GROUP BY r6
  order by r6;
  
DROP TABLE IF EXISTS sp_3; /* копирует в новую таблицу sp_3 все что в select */
CREATE TABLE sp_3(bok int,shap int);
insert into  sp_3(bok,shap)
  SELECT r6, count(at_0) FROM atoc_sp_1
  where r5=2
  GROUP BY r6
  order by r6;

DROP TABLE IF EXISTS sp_pr; /* копирует в новую таблицу roz_p_0 все что в select */
CREATE TABLE sp_pr(bok int,vsego int, mug int, gen int);
insert into  sp_pr(bok,vsego,mug,gen)
  SELECT t1.bok, t1.shap, t2.shap, t3.shap FROM sp_1 t1
  join sp_2 t2 ON t1.bok=t2.bok
  JOIN SP_3 t3 ON t2.BOK=t3.BOK
  GROUP BY t1.bok
  order by t1.bok;

DROP TABLE IF EXISTS sp_ok;/* копирует в новую таблицу sp_ok все что в select */
CREATE TABLE sp_ok(nac int,vsego int, muj int, jen int);
insert into  sp_ok(nac,vsego,muj,jen)
  SELECT t1.tx, t2.vsego, t2.mug, t2.gen FROM slv_sp_nac t1
  join sp_pr t2 ON t1.kd=t2.bok
  GROUP BY t1.KD
  order by t1.kd;

