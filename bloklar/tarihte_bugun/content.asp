<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Tarihte Bugun Blok Uygulamasi V 1.4										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
m1 = "Ocak"
m2 = "Þubat"
m3 = "Mart"
m4 = "Nisan"
m5 = "Mayýs"
m6 = "Haziran"
m7 = "Temmuz"
m8 = "Aðustos"
m9 = "Eylül"
m10 = "Ekim"
m11 = "Kasým"
m12 = "Aralýk"
ay = month(date)
strToddate = Day(Now())
strayi = Month(Now())
set tarihtebugun = Server.CreateObject("ADODB.Recordset")
tarihtebugun.ActiveConnection = Connection
tarihtebugun.Source = "SELECT olay, ayi, gunu, yili From tb WHERE ayi = " & strayi & " AND gunu = " & strToddate & ""
tarihtebugun.Open()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="td_menu">
<%
While NOT tarihtebugun.EOF
if tarihtebugun("yili") = "0" then
yil = ""
else
yil = tarihtebugun("yili")
end if
If ay = 1 Then
ayy = m1
elseIf ay = 2 Then
ayy = m2
elseIf ay = 3 Then
ayy = m3
elseIf ay = 4 Then
ayy = m4
elseIf ay = 5 Then
ayy = m5
elseIf ay = 6 Then
ayy = m6
elseIf ay = 7 Then
ayy = m7
elseIf ay = 8 Then
ayy = m8
elseIf ay = 9 Then
ayy = m9
elseIf ay = 10 Then
ayy = m10
elseIf ay = 11 Then
ayy = m11
elseIf ay = 12 Then
ayy = m12
End if
%>	
	<TR onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<TD><B><%=tarihtebugun("gunu")%>&nbsp;<%=ayy%>&nbsp;<%=yil%></B> : <%=tarihtebugun("olay")%></TD>
	</TR>
	<% 
	tarihtebugun.MoveNext
	Wend
	%>
</TABLE>
<%
tarihtebugun.Close
Set tarihtebugun = Nothing
Connection.close
set Connection=nothing
%>