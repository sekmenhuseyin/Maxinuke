<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
IF QS_CLEAR(Request.QueryString("action")) = "Member" Then
IF Session("Enter") = "Yes" then
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=uyemenu_uyeara%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<form action="?action=MemberSearch" method="post">
<table border="0" cellspacing="2" cellpadding="1" width="80%" align="center" class="td_menu" style="font-weight:bold">
<tr>
<td width="40%" align="right"><%=ms_criterion%> : </td>
<td width="60%" align="left"><input type="text" name="criterion" size="21" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right"><%=ms_select%> : </td>
<td width="60%" align="left">
<select name="stype" size="1" class="form1">
<option value="memname"><%=sby_memname%></option>
<option value="email"><%=sby_email%></option>
<option value="name"><%=sby_name%></option>
<option value="city"><%=sby_city%></option>
<option value="job"><%=sby_job%></option>
</select>
</tr>
<tr>
<td width="40%">&nbsp;</td>
<td width="60%"><input type="submit" value="<%=search_button%>" class="submit"></td>
</tr>
</table>
</form>
<%
ELSE
Response.Write no_entry
call ORTA2
call ALT
Response.End
End IF
ELSEIF QS_CLEAR(Request.QueryString("action")) = "MemberSearch" Then
crit = duz(Request.Form("criterion"))
stype = Request.Form("stype")

If stype = "memname" Then
stype = "kul_adi"
ElseIf stype = "email" Then
stype = "email"
ElseIf stype = "name" Then
stype = "isim"
ElseIf stype = "city" Then
stype = "sehir"
ElseIf stype = "job" Then
stype = "meslek"
End If
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=uyemenu_uyeara%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<%
If crit="" Then
Response.Write "<center><font face=verdana size=2>"&ms_error2&"</font></center>"
ELSE

Set mem = Server.CreateObject("ADODB.RecordSet")
mem.open "Select * FROM MEMBERS WHERE "&stype&" LIKE '%"&crit&"%'",Connection,3,3
If mem.Eof Then
Response.Write "<center><font face=verdana size=2>"&ms_error1&"</font></center>"
ELSE
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""2"" width=""100%"" align=""center"" style=""font-family:tahoma; font-size:11px;"" bgcolor="&tablo_cerceve&">"
Response.Write "<tr height=""18"" bgcolor="""&content_bg&""" style=""font-weight:bold""><td width=50 background=images/temalar/"&Session("SiteTheme")&"/menu_bg.gif>Avatar</td><td background=""images/temalar/"&Session("SiteTheme")&"/menu_bg.gif"">"&detail_memname&"</td><td background=images/temalar/"&Session("SiteTheme")&"/menu_bg.gif>"&detail_name&"</td></tr>"
Do While NOT mem.Eof
%>
	<tr bgcolor=<%=content_bg%>>
		<td>
		<%
		If mem("resmim_onay")=true Then
		resim=mem("resmim")
		Else
		resim="yok.gif"
		End if
		%>
		<IMG SRC="uploads/avatar/<%=resim%>" WIDTH="50" BORDER="0">
		</td>
		<td><a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><%=mem("kul_adi")%></a></td>
		<td><%=mem("isim")%></td>
	</tr>
<%
mem.MoveNext
Loop
Response.Write "</table>"
Response.Write "<br><br>"
Response.Write "<center><font face=verdana size=1><b>*</b> "&look_details&"</font></center>"
END IF
END IF

ELSE
kelam = duz(Request.QueryString("q"))
kelam = QS_CLEAR(kelam)
Set rs1 = Connection.Execute("Select * FROM NEWS WHERE baslik LIKE '%"&kelam&"%' OR metin LIKE '%"&kelam&"%'")
Set rs2 = Connection.Execute("Select * FROM DOWNLOADS WHERE p_name LIKE '%"&kelam&"%'")
Set rs3 = Connection.Execute("Select * FROM makale WHERE a_title LIKE '%"&kelam&"%'")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=arama%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<%
If kelam="" Then
Response.Write "<center><font class=td_menu><b>"&s_error0&"</b></font></center>"
ELSE
If rs1.Eof Then
Response.Write "<center><font class=td_menu><b>"&s_error1&"<b></font></center>"
ELSE %>
<div align="left" style="font-weight:bold"><%=top_menu2%> : </div>
<table border="0" cellspacing="1" cellpadding="1" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
<% do while not rs1.eof %>
<tr><td width="70%" bgcolor="<%=menu_color%>"><a href="news.asp?action=Read&hid=<%=rs1("hid")%>"><%=rs1("baslik")%></a> ( <%=rs1("ekleyen")%> ) </td><td align="center" bgcolor="<%=menu_color%>" width="30%"><%=rs1("tarih")%></td></tr>
<% rs1.MoveNext
Loop %>
</table>
<%
END IF
If rs2.Eof Then
Response.Write "<center><font class=td_menu><b>"&s_error2&"</b></font></center>"
ELSE %>
<hr size="1" width="100%" color="<%=tablo_cerceve%>">
<div align="left" style="font-weight:bold"><%=top_menu3%> : </div>
<table border="0" cellspacing="1" cellpadding="1" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
<% do while not rs2.eof %>
<tr><td width="70%" bgcolor="<%=menu_color%>"><a href="programs.asp?action=p_details&pid=<%=rs2("pid")%>"><%=rs2("p_name")%></a> ( <%=rs2("p_author")%> ) </td><td align="center" bgcolor="<%=menu_color%>" width="30%"><%=rs2("p_date")%></td></tr>
<% rs2.MoveNext
Loop %>
</table>
<%
END IF
If rs3.Eof Then
Response.Write "<center><font class=td_menu><b>"&s_error3&"</b></font></center>"
ELSE %>
<hr size="1" width="100%" color="<%=tablo_cerceve%>">
<div align="left" style="font-weight:bold"><%=top_menu4%> : </div>
<table border="0" cellspacing="1" cellpadding="1" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
<% do while not rs3.eof %>
	<tr>
		<td width="70%"><a href="makale.asp?action=Read&aid=<%=rs3("aid")%>"><%=rs3("a_title")%></a> ( <%=rs3("ekleyen")%> ) </td>
		<td align="center" bgcolor="<%=menu_color%>" width="30%"><%=rs3("a_date")%></td>
	</tr>
<%
rs3.MoveNext
Loop
%>
</table>
<%
END IF
END IF
END IF
%>








</td>
</tr>
</table>



<%
call ORTA2
call ALT
%>