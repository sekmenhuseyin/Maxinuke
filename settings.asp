<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
act = Request.QueryString("action")
	IF act = "settings" Then
	call settings
	ElseIF act = "update" Then
	call updSet
	End If

Sub settings
Set rs = Connection.Execute("Select * FROM SETTINGS")
%>
<form action="settings.asp?action=update" method="post">
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="4" align="center"><B>-=- GENEL SÝTE AYARLARI -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site Adý</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site Adresi</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site Slogani</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site Maili</B></td>
	</tr>
	<tr>
		<td><input type="text" name="name" value="<%=rs("site_isim")%>" class="form1"></td>
		<td><input type="text" name="url" value="<%=rs("site_adres")%>" class="form1"></td>
		<td><input type="text" name="site_slogan" value="<%=rs("site_slogan")%>" class="form1"></td>
		<td><input type="text" name="mail" value="<%=rs("site_mail")%>" class="form1"></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Üyelik Sistemi</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Þifre Hatýrlatma Sistemi</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site Ana Temasý</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site LCID</B></td>
	</tr>
	<tr>
		<td>
		<%
		If rs("onaylama") = True Then
		sec1 = "selected"
		ElseIf rs("onaylama") = False Then
		sec2 = "selected"
		End If
		%>
		<select name="onaylama" size="1" class="form1">
		<option value="False" <%=sec2%>>Direk Üyelik</option>
		<option value="True" <%=sec1%>>Mail Onaylý Üyelik</option>
		</select>
		</td>
		<td>
		<%
		If rs("fpass") = "Site" Then
		opt1 = "selected"
		ElseIf rs("fpass") = "Mail" Then
		opt2 = "selected"
		Else
		opt1 = "selected"
		End If
		%>
		<select name="fpass" size="1" class="form1">
		<option value="Site" <%=opt1%>>Site Üzerinden</option>
		<option value="Mail" <%=opt2%>>Mail Yolu Ýle
		</select>
		<td>
		<%
		Set themes = Server.CreateObject("ADODB.RecordSet")
		themes.open "Select * FROM THEMES ORDER BY t_name ASC",Connection,3,3
		%>
		<select name="theme" size="1" class="form1">
		<%
		Do Until themes.EoF
		IF rs("theme") = themes("t_dir") Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		Response.Write "<option value="""&themes("t_dir")&""" "&opt&">"&themes("t_name")&"</option>"
		themes.MoveNext
		Loop
		%>
		</select>
		</td>
		<td>
		<%
		If rs("s_lcid") = "1055" Then
		opt1 = "selected"
		ElseIf rs("s_lcid") = "1033" Then
		opt2 = "selected"
		ElseIf rs("s_lcid") = "2057" Then
		opt3 = "selected"
		ElseIf rs("s_lcid") = "1026" Then
		opt4 = "selected"
		ElseIf rs("s_lcid") = "1030" Then
		opt5 = "selected"
		ElseIf rs("s_lcid") = "1031" Then
		opt6 = "selected"
		ElseIf rs("s_lcid") = "1053" Then
		opt7 = "selected"
		End If
		%>
		<select name="s_lcid" size="1" class="form1">
		<option value="1055" <%=opt1%>>Türkiye</option>
		<option value="1033" <%=opt2%>>United States</option>
		<option value="2057" <%=opt3%>>United Kingdom</option>
		<option value="1026" <%=opt4%>>Bulgaristan</option>
		<option value="1030" <%=opt5%>>Danimarka</option>
		<option value="1031" <%=opt6%>>Almanya</option>
		<option value="1053" <%=opt7%>>Isvec</option>
		</select>
		</td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Site Logo Tipi</td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Site Logosu</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Bir Sayfada Listelenecek Program Sayýsý</B></td>	
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Mesaj Limiti</B></td>
	</tr>
	<tr>
		<td>
		<%
		If rs("logo_tipi")=false Then
		a="checked"
		elseIf rs("logo_tipi")=true Then
		b="checked"
		End if
		%>
		JPG,GIF,PNG,BMP <INPUT value=false TYPE="radio" NAME="logo_tipi" <%=a%>> Flash (SWF)<INPUT value=true TYPE="radio" NAME="logo_tipi" <%=b%>>  (Max:310X60)</td>
		<td><input type="text" name="site_logo" value="<%=rs("site_logo")%>" class="form1"></td>
		<td>
		<select name="prgs" class="form1" size="1">
		<%
		For I = 1 TO 50
		IF rs("prg_sayisi") = I Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<% Next %>
		</td>
		<td>
		<select name="msg_limit" class="form1" size="1">
		<%
		For I = 1 TO 1000
		IF rs("msg_limit") = I Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Forum Flood Koruma Süresi (Dakika)</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Forum Konu Sayisi</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Forum Cevap Sayisi</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Google Reklamlarý</B></td>
	</tr>
	<tr>
		<td>
		<select name="flood_time" size="1" class="form1">
		<%
		For I = 1 TO 60
		IF I = rs("flood_time") Then
		strOPT = "seLected"
		ELSE
		strOPT = ""
		End IF
		Response.Write "<option value="""&I&""" "&strOPT&">" & I & "</option>"
		Next
		%>
		</select>
		</td>
		<td>
		<select size="1" name="f_posts" class="form1">
		<%
		For I = 1 TO 50
		IF Rs("f_posts") = I Then
		opt = "selected"
		Else
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<% Next %>
		</select>
		</td>
		<td>
		<select size="1" name="f_replies" class="form1">
		<%
		For I = 1 TO 50
		IF Rs("f_replies") = I Then
		opt = "selected"
		Else
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<% Next %>
		</select>
		</td>
		<td>
		<select name="google" class="form1" size="1">
		<%
		IF rs("google") = True Then
		opt15 = "selected"
		ELSE
		opt16 = "selected"
		End IF
		%>
		<option value="True" <%=opt15%>>Açýk</option>
		<option value="False" <%=opt16%>>Kapalý</option>
		</td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Dakika Haberleri</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Dosyalar</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Makaleler</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son uyeler</B></td>
	</tr>
	<tr>
		<td>
		<%
		If rs("son_dakika") = true Then
		sddfg1 = "selected"
		ElseIf rs("son_dakika") = false Then
		sddfg2 = "selected"
		End If
		%>
		<select name="son_dakika" size="1" class="form1">
		<option value="true" <%=sddfg1%>>Olsun</option>
		<option value="false" <%=sddfg2%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("sd") = true Then
		zcv1 = "selected"
		ElseIf rs("sd") = false Then
		zcv2 = "selected"
		End If
		%>
		<select name="sd" size="1" class="form1">
		<option value="true" <%=zcv1%>>Olsun</option>
		<option value="false" <%=zcv2%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("sm") = true Then
		frty1 = "selected"
		ElseIf rs("sm") = false Then
		frty12 = "selected"
		End If
		%>
		<select name="sm" size="1" class="form1">
		<option value="true" <%=frty1%>>Olsun</option>
		<option value="false" <%=frty2%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("su") = true Then
		futr1 = "selected"
		ElseIf rs("su") = false Then
		futr2 = "selected"
		End If
		%>
		<select name="su" size="1" class="form1">
		<option value="true" <%=futr1%>>Olsun</option>
		<option value="false" <%=futr2%>>Olmasin</option>
		</select>
		</td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Fikralar</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Ekartlar</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Forumlar</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Resimler</B></td>
	</tr>
	<tr>
		<td>
		<%
		If rs("sf") = true Then
		sektor1 = "selected"
		ElseIf rs("sf") = false Then
		sektor2 = "selected"
		End If
		%>
		<select name="sf" size="1" class="form1">
		<option value="true" <%=sektor1%>>Olsun</option>
		<option value="false" <%=sektor2%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("se") = true Then
		opthh1 = "selected"
		ElseIf rs("se") = false Then
		opthh2 = "selected"
		End If
		%>
		<select name="se" size="1" class="form1">
		<option value="true" <%=opthh1%>>Olsun</option>
		<option value="false" <%=opthh2%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("sfb") = true Then
		sektor3 = "selected"
		ElseIf rs("sfb") = false Then
		sektor4 = "selected"
		End If
		%>
		<select name="sfb" size="1" class="form1">
		<option value="true" <%=sektor3%>>Olsun</option>
		<option value="false" <%=sektor4%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("sr") = true Then
		sektor5 = "selected"
		ElseIf rs("sr") = false Then
		sektor6 = "selected"
		End If
		%>
		<select name="sr" size="1" class="form1">
		<option value="true" <%=sektor5%>>Olsun</option>
		<option value="false" <%=sektor6%>>Olmasin</option>
		</select>
		</td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Oyunlar</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfada Son Videolar</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Radio MP3 Player</B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><B>Ana Sayfa Alt Haberler</B></td>
	</tr>
	<tr>
		<td>
		<%
		If rs("so") = true Then
		sektor7 = "selected"
		ElseIf rs("so") = false Then
		sektor8 = "selected"
		End If
		%>
		<select name="so" size="1" class="form1">
		<option value="true" <%=sektor7%>>Olsun</option>
		<option value="false" <%=sektor8%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("sv") = true Then
		sektor9 = "selected"
		ElseIf rs("sv") = false Then
		sektor10 = "selected"
		End If
		%>
		<select name="sv" size="1" class="form1">
		<option value="true" <%=sektor9%>>Olsun</option>
		<option value="false" <%=sektor10%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("rplayer") = true Then
		sektor11 = "selected"
		ElseIf rs("rplayer") = false Then
		sektor12 = "selected"
		End If
		%>
		<select name="rplayer" size="1" class="form1">
		<option value="true" <%=sektor11%>>Olsun</option>
		<option value="false" <%=sektor12%>>Olmasin</option>
		</select>
		</td>
		<td>
		<%
		If rs("hv") = true Then
		sektorhv1 = "selected"
		ElseIf rs("hv") = false Then
		sektorhv2 = "selected"
		End If
		%>
		<select name="hv" size="1" class="form1">
		<option value="true" <%=sektorhv1%>>Olsun</option>
		<option value="false" <%=sektorhv2%>>Olmasin</option>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan=4 background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><CENTER>KEYWORDS</CENTER></td>
	</tr>
	<tr>
		<td colspan=4 ><CENTER><textarea name="Keywords" style="width:100%" rows="5" class="form1"><%=rs("Keywords")%></textarea></CENTER></td>
	</tr>
	<tr>
		<td colspan=4 background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><CENTER>Üyeler Kapalý Sayfada Çýkacak Yazý</CENTER></td>
	</tr>
	<tr>
		<td colspan=4 ><CENTER><textarea name="npm" style="width:100%" rows="10" class="form1"><%=rs("np_msg")%></textarea></CENTER></td>
	</tr>
	<tr>
		<td colspan=4 background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><CENTER>Uyelik Sozlesmesi</CENTER></td>
	</tr>
	<tr>
		<td colspan=4 ><CENTER><textarea name="sozlesme" style="width:100%" rows="6" class="form1"><%=rs("sozlesme")%></textarea></CENTER></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" value="Deðiþiklikleri Kaydet" class="submit" style="width:100%"></td>
	</tr>
</table>
</form>
<%
themes.Close
Set themes = Nothing
rs.Close
Set rs = Nothing
End Sub

Sub updSet
sname = duz(Request.Form("name"))
surl = duz(Request.Form("url"))
smail = duz(Request.Form("mail"))
sprgs = duz(Request.Form("prgs"))
sml = duz(Request.Form("msg_limit"))
snpm = Request.Form("npm")
sozlesme = Request.Form("sozlesme")
stheme = duz(Request.Form("theme"))
sfpass = duz(Request.Form("fpass"))
s_lcid = duz(Request.Form("s_lcid"))
son_dakika = duz(Request.Form("son_dakika"))
site_slogan = duz(Request.Form("site_slogan"))
logo_tipi = duz(Request.Form("logo_tipi"))
site_logo = duz(Request.Form("site_logo"))
onaylama = duz(Request.Form("onaylama"))
Keywords = duz(Request.Form("Keywords"))
hv = Request.Form("hv")
If sname="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Site Ýsmini Yazýnýz...</center>"
ElseIf surl="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Site Adresini Yazýnýz...</center>"
ElseIf smail="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Site Mailini Yazýnýz...</center>"
ElseIf sprgs="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Bir Sayfada Listelenecek Program Sayýsýný Yazýnýz...</center>"
ElseIf sml="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Mesaj Limitini Yazýnýz...</center>"
ElseIf snpm="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Üyeler Kapalý Sayfada çýkacak Yazýyý Yazýnýz...</center>"
ELSE
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM SETTINGS"
ent.open entSQL,Connection,1,3
ent("site_isim") = sname
ent("site_logo") = site_logo
ent("site_adres") = surl
ent("site_mail") = smail
ent("prg_sayisi") = sprgs
ent("msg_limit") = sml
ent("np_msg") = snpm
ent("sozlesme") = sozlesme
ent("theme") = stheme
ent("fpass") = sfpass
ent("s_lcid") = s_lcid
ent("son_dakika") = son_dakika
ent("f_posts") = Request.Form("f_posts")
ent("f_replies") = Request.Form("f_replies")
ent("site_slogan") = Request.Form("site_slogan")
ent("flood_time") = Request.Form("flood_time")
ent("google") = Request.Form("google")
ent("onaylama") = Request.Form("onaylama")
ent("logo_tipi") = Request.Form("logo_tipi")
ent("Keywords") = Request.Form("Keywords")
ent("sd") = Request.Form("sd")
ent("sm") = Request.Form("sm")
ent("su") = Request.Form("su")
ent("sf") = Request.Form("sf")
ent("se") = Request.Form("se")
ent("sfb") = Request.Form("sfb")
ent("sr") = Request.Form("sr")
ent("so") = Request.Form("so")
ent("sv") = Request.Form("sv")
ent("rplayer") = Request.Form("rplayer")
ent("hv") = Request.Form("hv")
ent.Update
%>
<script>alert('Guncelleme Basarili');location.href='settings.asp?action=settings';</script>
<%
ent.Close
Set ent = Nothing
END IF
End Sub
End If
%>