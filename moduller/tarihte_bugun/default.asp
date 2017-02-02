<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Tarihte Bugun Modul Uygulamasi V 1.0										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
strToddate = Day(Now())
strayi = Month(Now())
set tarihtebugun = Server.CreateObject("ADODB.Recordset")
tarihtebugun.ActiveConnection = Connection
tarihtebugun.Source = "SELECT olay, ayi, gunu, yili From tb WHERE ayi = " & strayi & " AND gunu = " & strToddate & ""
tarihtebugun.Open()
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<TR height="20">
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" align="center"><b>TARÝH</TD>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" align="center"><b>OLAY</TD>
	</TR>
<%
While NOT tarihtebugun.EOF
if tarihtebugun("yili") = "0" then
yil = ""
else
yil = tarihtebugun("yili")
end if
%>	
	<TR onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<TD valign="top"><%=tarihtebugun("gunu")%>&nbsp;<%=MonthName(Month(Date()))%>&nbsp;<%=yil%></TD>
		<TD><%=tarihtebugun("olay")%></TD>
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