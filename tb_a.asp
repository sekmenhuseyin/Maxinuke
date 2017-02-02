<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=content_bg%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
IF act = "hepsi" Then
call hepsii
ElseIf act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "new_enter" Then
call enternew
ElseIf act = "register" Then
call reg
ElseIf act = "delete" Then
call del
End If

Sub hepsii
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM tb ORDER BY hid DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Olay YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM Olaylar -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Olay</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Yýlý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Günü</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ayý</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
%>
	<tr>
		<td align="center"><%=rs("hid")%></td>
		<td><%=rs("olay")%></td>
		<td align="center"><%=rs("yili")%></td>
		<td align="center"><%=rs("gunu")%></td>		
		<td align="center"><%=rs("ayi")%></td>				
		<td align="center"><a href="?action=organize&hid=<%=rs("hid")%>">Düzenle</a> - <a href="?action=delete&hid=<%=rs("hid")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub

Sub enternew
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.ayi.value == "0") {alert("Lütfen Olay Ayini Seciniz... !"); return false; }
if (form.gunu.value == "0") {alert("Lütfen Olay Gununu Seciniz... !"); return false; }
if (form.olay.value == "") {alert("Lütfen Olay Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width="100%">
	<tr>
		<td>Yili</td>
		<td>
		<select name="yili" size="1" class="form1">
		<% For i = 0 TO 2009 %>
		<option value="<%=i%>"><%=i%></option>"
		<% Next %>
		</select>
		</td>
		<td>Ayi</td>
		<td>
		<select name="ayi" size="1" class="form1">
		<% For i = 0 TO 12 %>
		<option value="<%=i%>"><%=i%></option>"
		<% Next %>
		</select>
		</td>
		<td>Gunu</td>
		<td>
		<select name="gunu" size="1" class="form1">
		<% For i = 0 TO 31 %>
		<option value="<%=i%>"><%=i%></option>"
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="6">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= ""
oFCKeditor.Create "olay"
%>	</td>
	</tr>
	<tr>
		<td align="center" colspan="6"><input type="submit" class="submit"  name="submit" style="width:100%" value="Olayi Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg
yili = Request("yili")
ayi = Request("ayi")
gunu = Request("gunu")
olay = Request("olay")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM tb",Connection,3,3
ent.AddNew
ent("yili") = yili
ent("ayi") = ayi
ent("gunu") = gunu
ent("olay") = olay
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=hepsi"
End Sub

Sub organize
id = Request.QueryString("hid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM tb WHERE hid = "&id&""
rs.open SQL,Connection,1,3

%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.ayi.value == "0") {alert("Lütfen Olay Ayini Seciniz... !"); return false; }
if (form.gunu.value == "0") {alert("Lütfen Olay Gununu Seciniz... !"); return false; }
if (form.olay.value == "") {alert("Lütfen Olay Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&hid=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Yili</td>
		<td>
		<select name="yili" size="1" class="form1">
		<% For i = 0 TO 2009 %>
		<option value="<%=i%>" <%if rs("yili") = i Then response.write "selected" end if%>><%=i%></option>"
		<% Next %>
		</select>
		</td>
		<td>Ayi</td>
		<td>
		<select name="ayi" size="1" class="form1">
		<% For i = 0 TO 12 %>
		<option value="<%=i%>" <%if rs("ayi") = i Then response.write "selected" end if%>><%=i%></option>
		<% Next %>
		</select>
		</td>
		<td>Gunu</td>
		<td>
		<select name="gunu" size="1" class="form1">
		<% For i = 0 TO 31 %>
		<option value="<%=i%>" <%if rs("gunu") = i Then response.write "selected" end if%>><%=i%></option>"
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="6">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= rs("olay")
oFCKeditor.Create "olay"
%></td>
	</tr>
	<tr>
		<td align="center" colspan="6"><input type="submit" class="submit"  name="submit" style="width:100%" value="olayi Güncelle"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

id = Request.QueryString("hid")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM tb WHERE hid = "&id&""
ent.open entSQL,Connection,1,3
yili = Request.Form("yili")
ayi = Request.Form("ayi")
gunu = Request.Form("gunu")
olay = Request.Form("olay")
ent("yili") = yili
ent("ayi") = ayi
ent("gunu") = gunu
ent("olay") = olay
ent.Update
Response.Redirect "?action=hepsi"
ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("hid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM tb WHERE hid = "&id&""
delete.open delSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
%>