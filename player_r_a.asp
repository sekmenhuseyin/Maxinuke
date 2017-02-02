<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "new_enter" Then
call enternew
ElseIf act = "register" Then
call reg
ElseIf act = "delete" Then
call del
ElseIf act = "change" Then
call change
ElseIF act = "AllNews" Then
call all_news
End If

Sub all_news
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM muzikler where tip=2 ORDER BY sanatci",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Muzik YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM RADYOLAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Radyo Adi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Radyo Slogani</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Radyo Adresi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
%>
	<tr>
		<td align="center"><%=rs("muzik_id")%></td>
		<td><%=rs("sanatci")%></td>
		<td><%=rs("parca")%></td>
		<td><%=rs("dosya_adi")%></td>
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&muzik_id=<%=rs("muzik_id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&muzik_id=<%=rs("muzik_id")%>">Pasifleþtir</a> <%End if%>- <a href="?action=organize&muzik_id=<%=rs("muzik_id")%>">Düzenle</a> - <a href="?action=delete&muzik_id=<%=rs("muzik_id")%>">Sil</a></td>
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
if (form.sanatci.value == "") {alert("Lütfen Radyo Istasyonun adini yaziniz... !"); return false; }
if (form.parca.value == "") {alert("Lütfen Radyo Istasyonun sloganini Yazýnýz... !"); return false; }
if (form.dosya_adi.value == "") {alert("Lütfen muzik dosyasini bilgisayarinizdan yukleyiniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Radyo Adi</td>
		<td>: <input type="text" name="sanatci" class="form1" size="60"></td>
	</tr>
	<tr>
		<td>Radyo Slogani</td>
		<td>: <input type="text" name="parca" class="form1" size="60"></td>
	</tr>
	<tr>
		<td>Radyo Adresi</td>
		<td>: <input type="text" name="dosya_adi" class="form1" size="60"></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Radyoyu Ekle"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg
sanatci = Request("sanatci")
parca = Request("parca")
dosya_adi = Request("dosya_adi")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM muzikler",Connection,3,3
ent.AddNew
ent("sanatci") = sanatci
ent("parca") = parca
ent("dosya_adi") = dosya_adi
ent("onay") = True
ent("tip") = "2"
ent.Update
ent.Close
Set ent = Nothing
session("dosya_adi")=""
Response.Redirect "?action=AllNews"
End Sub
Sub change

operation = Request.QueryString("op")
id = Request.QueryString("muzik_id")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM muzikler WHERE muzik_id = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")

chn.Close
Set chn = Nothing

End Sub
Sub organize

id = Request.QueryString("muzik_id")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM muzikler WHERE muzik_id = "&id&""
rs.open SQL,Connection,1,3
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.sanatci.value == "") {alert("Lütfen Radyo Istasyonun adini yaziniz... !"); return false; }
if (form.parca.value == "") {alert("Lütfen Radyo Istasyonun sloganini Yazýnýz... !"); return false; }
if (form.dosya_adi.value == "") {alert("Lütfen muzik dosyasini bilgisayarinizdan yukleyiniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&muzik_id=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Radyo Adi</td>
		<td>: <input type="text" name="sanatci" class="form1" size="60" value="<%=rs("sanatci")%>"></td>
	</tr>
	<tr>		
		<td>Radyo Slogani</td>
		<td>: <input type="text" name="parca" class="form1" size="60" value="<%=rs("parca")%>"></td>
	</tr>
	<tr>		
		<td>Radyo Adresi</td>
		<td>: <input type="text" name="dosya_adi" class="form1" size="60" value="<%=rs("dosya_adi")%>"></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Radyoyu Güncelle"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

id = Request.QueryString("muzik_id")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM muzikler WHERE muzik_id = "&id&""
ent.open entSQL,Connection,1,3
sanatci = Request.Form("sanatci")
parca = Request.Form("parca")
dosya_adi = Request.Form("dosya_adi")
ent("sanatci") = sanatci
ent("parca") = parca
ent("dosya_adi") = dosya_adi
ent.Update
Response.Redirect "?action=AllNews"
ent.Close
Set ent = Nothing
End Sub

Sub del
id = Request.QueryString("muzik_id")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM muzikler WHERE muzik_id = "&id&""
delete.open delSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>