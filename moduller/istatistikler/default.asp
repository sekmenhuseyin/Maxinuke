<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Istatistikler Modul Uygulamasi											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
Set makale_onayli = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved = True")
Set makale_onaysiz = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved = False")
Set makale_okunma = Connection.Execute("SELECT SUM(hit) AS count FROM ARTICLES")
Set makale_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'articles'")
Set dosya_onayli = Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE onay = True")
Set dosya_onaysiz = Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE onay = False")
Set dosya_indirilme = Connection.Execute("SELECT SUM(p_hit) AS count FROM DOWNLOADS")
Set dosya_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'programs'")
Set ekart_onayli = Connection.Execute("Select Count(*) AS count FROM ekartlar WHERE onay = True")
Set ekart_onaysiz = Connection.Execute("Select Count(*) AS count FROM ekartlar WHERE onay = False")
Set ekart_gonderilme = Connection.Execute("SELECT SUM(hit) AS count FROM ekartlar")
Set fikra_onayli = Connection.Execute("Select Count(*) AS count FROM fikra WHERE onay = True")
Set fikra_onaysiz = Connection.Execute("Select Count(*) AS count FROM fikra WHERE onay = False")
Set fikra_okunma = Connection.Execute("SELECT SUM(hit) AS count FROM fikra")
Set fikra_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'fikra'")
Set fo_onayli = Connection.Execute("Select Count(*) AS count FROM fo WHERE onay = True")
Set fo_onaysiz = Connection.Execute("Select Count(*) AS count FROM fo WHERE onay = False")
Set fo_oynanma = Connection.Execute("SELECT SUM(oynanma) AS count FROM fo")
Set fo_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'fo'")
Set zd_onayli = Connection.Execute("Select Count(*) AS count FROM GUESTBOOK WHERE onay = True")
Set zd_onaysiz = Connection.Execute("Select Count(*) AS count FROM GUESTBOOK WHERE onay = False")
Set aboneler = Connection.Execute("Select Count(*) AS count FROM f_abone")
Set arkadaslar = Connection.Execute("Select Count(*) AS count FROM FRIENDS")
Set admin_genel = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 1")
Set admin_dosya = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 2")
Set admin_makale = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 3")
Set admin_haber = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 4")
Set admin_forum = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 5")
Set icmesaj = Connection.Execute("Select Count(*) AS count FROM MESSAGES")
Set haber_onayli = Connection.Execute("Select Count(*) AS count FROM NEWS WHERE onay = True")
Set haber_onaysiz = Connection.Execute("Select Count(*) AS count FROM NEWS WHERE onay = False")
Set haber_okunma = Connection.Execute("SELECT SUM(okunma) AS count FROM NEWS")
Set haber_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'NEWS'")
Set rg_onayli = Connection.Execute("Select Count(*) AS count FROM rg WHERE onay = True")
Set rg_onaysiz = Connection.Execute("Select Count(*) AS count FROM rg WHERE onay = False")
Set rg_okunma = Connection.Execute("SELECT SUM(bakilma) AS count FROM rg")
Set rg_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'rg'")
Set uye_onayli = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE onay = True")
Set uye_onaysiz = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE onay = False")
Set toplam_giris = Connection.Execute("SELECT SUM(login_sayisi) AS count FROM MEMBERS")
Set video_onayli = Connection.Execute("Select Count(*) AS count FROM videolar WHERE onay = True")
Set video_onaysiz = Connection.Execute("Select Count(*) AS count FROM videolar WHERE onay = False")
Set video_okunma = Connection.Execute("SELECT SUM(hit) AS count FROM videolar")
Set video_yorum = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'videolar'")
Set temalar = Connection.Execute("Select Count(*) AS count FROM THEMES")
%>
<table border="0" cellspacing="3" cellpadding="3" width="100%" class="td_menu" >
	<tr>
		<td style="<%=TableBorder%>"><B>Makele : </B><B><%=makale_onayli("count")%></B> onaylý <B><%=makale_onaysiz("count")%></B> onaysýz toplam <B><%=makale_onayli("count")+makale_onaysiz("count")%></B> makale <% IF Len(makale_okunma("count")) >= 1 Then %><B><%=makale_okunma("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa okundu <B><%=makale_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Dosya : </B><B><%=dosya_onayli("count")%></B> onaylý <B><%=dosya_onaysiz("count")%></B> onaysýz toplam <B><%=dosya_onayli("count")+dosya_onaysiz("count")%></B> dosya <% IF Len(dosya_indirilme("count")) >= 1 Then %><B><%=dosya_indirilme("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa indirildi <B><%=dosya_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>E-Kart : </B><B><%=ekart_onayli("count")%></B> onaylý <B><%=ekart_onaysiz("count")%></B> onaysýz toplam <B><%=ekart_onayli("count")+ekart_onaysiz("count")%></B> Ekart <% IF Len(ekart_gonderilme("count")) >= 1 Then %><B><%=ekart_gonderilme("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa gönderildi.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Fýkra : </B><B><%=fikra_onayli("count")%></B> onaylý <B><%=fikra_onaysiz("count")%></B> onaysýz toplam <B><%=fikra_onayli("count")+fikra_onaysiz("count")%></B> fýkra <% IF Len(fikra_okunma("count")) >= 1 Then %><B><%=fikra_okunma("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa okundu <B><%=fikra_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Flaþh Oyun : </B><B><%=fo_onayli("count")%></B> onaylý <B><%=fo_onaysiz("count")%></B> onaysýz toplam <B><%=fo_onayli("count")+fo_onaysiz("count")%></B> Flash Oyun <% IF Len(fo_oynanma("count")) >= 1 Then %><B><%=fo_oynanma("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa oynandý <B><%=fo_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Ziyaretçi Defteri : </B><B><%=zd_onayli("count")%></B> onaylý <B><%=zd_onaysiz("count")%></B> onaysýz toplam <B><%=zd_onayli("count")+zd_onaysiz("count")%></B> ziyaretçi görüþü.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Yönetim : </B><B><%=admin_genel("count")%></B> Genel Yönetici, <B><%=admin_haber("count")%></B> Haber Editörü, <B><%=admin_dosya("count")%></B> Download Editörü, <B><%=admin_makale("count")%></B> Makale Editörü, <B><%=admin_forum("count")%></B> Forum Görevlisi.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Haber : </B><B><%=haber_onayli("count")%></B> onaylý <B><%=haber_onaysiz("count")%></B> onaysýz toplam <B><%=haber_onayli("count")+haber_onaysiz("count")%></B> Haber <% IF Len(haber_okunma("count")) >= 1 Then %><B><%=haber_okunma("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa okundu <B><%=haber_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Resim Galerisi : </B><B><%=rg_onayli("count")%></B> onaylý <B><%=rg_onaysiz("count")%></B> onaysýz toplam <B><%=rg_onayli("count")+rg_onaysiz("count")%></B> resim <% IF Len(rg_okunma("count")) >= 1 Then %><B><%=rg_okunma("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa görüntülendi <B><%=rg_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>"><B>Üyelik : </B><B>1<%=uye_onayli("count")%></B> onaylý <B><%=uye_onaysiz("count")%></B> onaysýz toplam <B>1<%=uye_onayli("count")+uye_onaysiz("count")%></B> üye var ve <B><%=toplam_giris("count")%></B> defa siteye login oldular.</td>
	</tr>
	<tr>
		<td <%=arka_v%> style="<%=TableBorder%>"><B>Videolar : </B><B><%=video_onayli("count")%></B> onaylý <B><%=video_onaysiz("count")%></B> onaysýz toplam <B><%=video_onayli("count")+video_onaysiz("count")%></B> video <% IF Len(video_okunma("count")) >= 1 Then %><B><%=video_okunma("count")%></B><%ELSE%><%Response.Write "<B>0</B>"%><% End IF %> defa izlendi <B><%=video_yorum("count")%></B> defa yorumlandý.</td>
	</tr>
	<tr>
		<td style="<%=TableBorder%>" bgcolor="<%=content_bg%>"><%=sett("site_isim")%> genelinde <B><%=temalar("count")%></B> tema, <B><%=aboneler("count")%></B> üyenin forumlarda aboneliði var ayrýca <B><%=arkadaslar("count")%></B> üye birbiriyle arkadaþlýk kurdu ve <B><%=icmesaj("count")%></B> mesaj yollandý.</Td>
	</tr>
</table>

<%
makale_onayli.Close
Set makale_onayli = Nothing
makale_onaysiz.Close
Set makale_onaysiz = Nothing
makale_okunma.Close
Set makale_okunma = Nothing
makale_yorum.Close
Set makale_yorum = Nothing
dosya_onayli.Close
Set dosya_onayli = Nothing
dosya_onaysiz.Close
Set dosya_onaysiz = Nothing
dosya_indirilme.Close
Set dosya_indirilme = Nothing
dosya_yorum.Close
Set dosya_yorum = Nothing
ekart_onayli.Close
Set ekart_onayli = Nothing
ekart_onaysiz.Close
Set ekart_onaysiz = Nothing
ekart_gonderilme.Close
Set ekart_gonderilme = Nothing
fikra_onayli.Close
Set fikra_onayli = Nothing
fikra_onaysiz.Close
Set fikra_onaysiz = Nothing
fikra_okunma.Close
Set fikra_okunma = Nothing
fikra_yorum.Close
Set fikra_yorum = Nothing
fo_onayli.Close
Set fo_onayli = Nothing
fo_onaysiz.Close
Set fo_onaysiz = Nothing
fo_oynanma.Close
Set fo_oynanma = Nothing
fo_yorum.Close
Set fo_yorum = Nothing
zd_onayli.Close
Set zd_onayli = Nothing
zd_onaysiz.Close
Set zd_onaysiz = Nothing
aboneler.Close
Set aboneler = Nothing
arkadaslar.Close
Set arkadaslar = Nothing
admin_genel.Close
Set admin_genel = Nothing
admin_dosya.Close
Set admin_dosya = Nothing
admin_makale.Close
Set admin_makale = Nothing
admin_haber.Close
Set admin_haber = Nothing
admin_forum.Close
Set admin_forum = Nothing
icmesaj.Close
Set icmesaj = Nothing
haber_onayli.Close
Set haber_onayli = Nothing
haber_onaysiz.Close
Set haber_onaysiz = Nothing
haber_okunma.Close
Set haber_okunma = Nothing
haber_yorum.Close
Set haber_yorum = Nothing
rg_onayli.Close
Set rg_onayli = Nothing
rg_onaysiz.Close
Set rg_onaysiz = Nothing
rg_okunma.Close
Set rg_okunma = Nothing
rg_yorum.Close
Set rg_yorum = Nothing
uye_onayli.Close
Set uye_onayli = Nothing
uye_onaysiz.Close
Set uye_onaysiz = Nothing
toplam_giris.Close
Set toplam_giris = Nothing
video_onayli.Close
Set video_onayli = Nothing
video_onaysiz.Close
Set video_onaysiz = Nothing
video_okunma.Close
Set video_okunma = Nothing
video_yorum.Close
Set video_yorum = Nothing
temalar.Close
Set temalar = Nothing
Connection.close
set Connection=nothing
%>