<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Max Giris Blok Uygulamasi V 1.0											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%Set maxgiris = Connection.Execute("Select * FROM WEBCOUNTER ORDER BY today DESC")%>
<TABLE class="td_menu" width="100%">
	<TR>
		<TD align="center">Bu site �imdiye kadar en fazla <B><%=maxgiris("cdate")%></B> tarihinde<BR><font color="red" size="7"><B><%=maxgiris("today")%></B></font><BR>ki�iyi a��rlad�.</TD>
	</TR>
</TABLE>
<%
maxgiris.Close
Set maxgiris = Nothing
%>