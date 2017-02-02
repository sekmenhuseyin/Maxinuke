<table background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" width="100%" border="0" cellspacing="0" cellpadding="0"  class="copy">
	<tr>
		<td align="center" height="20"><B>Copyright 2008 - 2009 © <%=sett("site_isim")%>&nbsp;-&nbsp;Powered By <a href="http://www.maxinuke.com" target="_blank">Maxi Nuke (V<%=v%>)</a> - <%=en%> - Her Hakký Saklýdýr. Sayfan&#305;z&#305;n yüklenme süresi: 
<%
lngTimer=Timer
For lngCnt=1 to 1000000
Next
Response.Write FormatNumber(Timer-lngTimer,2,True) & " saniye" 



IF Session("Enter") = "Yes" Then
uyebilgi.Close
Set uyebilgi = Nothing
End if


Connection.Close
Set Connection = Nothing
%></B>

		</td>

	</tr>
</table>
<% End Sub 'ALT %>
</body>
</html>