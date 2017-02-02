<%Sub ORTA%>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%" align="left" valign="top"> 
<%
Do Until LB.Eof
IF LB("b_active") = True Then
%>
<table border="0" cellpadding="0" cellspacing="0" width="165" class="td_menu">
	<TR>
		<TD colspan="3" background="images/temalar/<%=Session("SiteTheme")%>/ust.gif" height="21" align="center"><B><%=LB("b_name")%></B></TD>
	</TR>
	<TR>
		<td background="images/temalar/<%=Session("SiteTheme")%>/sol.gif" WIDTH="12"></td>
		<TD bgcolor="<%=content_bg%>" width="141">
		<%
		IF LB("b_include") <> "null" Then
		Server.Execute ("bloklar/"&LB("b_include")&"/content.asp")
		End IF
		Response.Write LB("b_content")
		%>
		</TD>
		<td background="images/temalar/<%=Session("SiteTheme")%>/sag.gif" WIDTH="12"></td>
	</TR>
	<TR>
		<td colspan="3"><img src="images/temalar/<%=Session("SiteTheme")%>/alt.gif"></td>
	</TR>
</TABLE>
<BR>
<%
End IF
LB.MoveNext
Loop
%>
    </td>
    <td width="60%" align="center" valign="top">
<%
End Sub 'ORTA
Sub ORTA2
%>
</td>
    <td width="20%" align="right" valign="top">
<%
Do Until RB.Eof
IF RB("b_active") = True Then
%>




<table border="0" cellpadding="0" cellspacing="0" width="165" class="td_menu">
	<TR>
		<TD colspan="3" background="images/temalar/<%=Session("SiteTheme")%>/ust.gif" height="21" align="center"><B><%=RB("b_name")%></B></TD>
	</TR>
	<TR>
		<td background="images/temalar/<%=Session("SiteTheme")%>/sol.gif" WIDTH="12"></td>
		<TD bgcolor="<%=content_bg%>" width="141">
		<%
		IF RB("b_include") <> "null" Then
		Server.Execute ("bloklar/"&RB("b_include")&"/content.asp")
		End IF
		Response.Write RB("b_content")
		%>
		</TD>
		<td background="images/temalar/<%=Session("SiteTheme")%>/sag.gif" WIDTH="12"></td>
	</TR>
	<TR>
		<td colspan="3"><img src="images/temalar/<%=Session("SiteTheme")%>/alt.gif"></td>
	</TR>
</TABLE>
<BR>
<%
End IF
RB.MoveNext
Loop
%>

    </td>
  </tr>
</table>
<%
End Sub 'ORTA2
Sub ALT
%>
<br>
<br>
<CENTER><!--#include file="reklam_alt.asp" --></CENTER>
<br>
<!--#include file="alt.asp" -->