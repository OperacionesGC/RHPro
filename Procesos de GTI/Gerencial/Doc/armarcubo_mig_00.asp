<% Option Explicit %>
<!--#include virtual="/rhprox2/shared/inc/sec.inc"-->
<!--#include virtual="/rhprox2/shared/inc/const.inc"-->
<!--#include virtual="/rhprox2/shared/db/conn_db.inc"-->
<!--#include virtual="/rhprox2/shared/inc/fecha.inc"-->
<!--#include virtual="/rhprox2/shared/inc/sqls.inc"-->
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
window.resizeTo(150,150);</script>
<% 
const l_valornulo = "null"

Dim l_anrcabnro

Dim l_cm
Dim l_ins
Dim l_rs
Dim l_sql


Dim l_hora
Dim l_arrhr

Dim l_codproc

Dim l_id
l_id = Session("Username")

'------------------------------------------	 funciones -------------------------------------

function strto2(cad)
	if cad<10 then
		strto2= "0" & cad
	else
		strto2= cad
	end if 
	
end function


' body --------------------------------------------------------------------------------

l_hora = mid(time,2,8)
'l_arrhr= Split(l_hora,".")
l_arrhr= Split(l_hora,":")

'Response.Write("<br>")
'Response.Write(l_arrhr)

l_hora = strto2(l_arrhr(0))&":"&l_arrhr(1)&":"&l_arrhr(2)

l_anrcabnro = Request.QueryString("anrcabnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn

l_sql = "insert into batch_proceso "
l_sql = l_sql & "(btprcnro, "
l_sql = l_sql & " bprcfecha, iduser, bprchora, "
l_sql = l_sql & " bprcfecdesde, bprcfechasta, "
l_sql = l_sql & " bprcparam, bprcestado, bprcprogreso, bprcfecfin, bprchorafin, "
l_sql = l_sql & " bprctiempo, empnro, bprcempleados) "
l_sql = l_sql & " VALUES (18," 
l_sql = l_sql &   cambiafecha(Date,"YMD",true) & ", '"& l_id &"','"& l_hora &"' "
l_sql = l_sql &   ", null, null, '" 
l_sql = l_sql &   l_anrcabnro & "', 'Pendiente', null , null, null, "
l_sql = l_sql &   " null, 0, null)"
'Response.Write(l_sql)
'Response.Write("<br>")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Proceso generado');window.close();</script>"
%>
