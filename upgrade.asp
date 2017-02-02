<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Not Session("Level") = "1" Then
response.redirect "default.asp"
Else
	If Request.QueryString("eylem") = "" then
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="center">Eðer MaxiNuke yada Mini-nuke´nin geçmiþ sürümlerinden birisini kullanýyorsanýz veritabanlarýnýzdaki bilgileri Maxi Nuke V <%=v%> e çekebiirsiniz.<BR>Kullandiginiz hazir portal mini yada maxinuke deðil ise veritabaninizýn maxinukeye uyarlanmasý icin <A HREF="http://www.maxinuke.com/moduller.asp?mi=18"><B>MAXINUKE ILETISIM SAYFASI´</B></A>NI kullanabilir yada direk info@maxinuke.com adresine veritaban dosyanýzý yollayabilirsiniz.</td>
	</tr>
	<tr>
		<td align="center"><a href="?eylem=sor&tip=maxinuke"><button onclick="parentElement.click();" CLASS="submit">>> MAXI-NUKE VERÝTABANINDAN BÝLGÝ AL >></button></a></td>
	</tr>
	<tr>
		<td align="center"><a href="?eylem=sor&tip=mininuke"><button onclick="parentElement.click();" CLASS="submit">>> MÝNÝ-NUKE VERÝTABANINDAN BÝLGÝ AL >></button></a></td>
	</tr>
</table>
<%
	elseIf Request.QueryString("eylem") = "sor" Then
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.eski_mdb.value == "") {alert("Bilgilerin alinacagi eski veritabaninin ismini yazmadiniz !"); return false; }
return true;
}
</SCRIPT>
<FORM onSubmit="return formkontrol(this)" name=peno action="?eylem=basla&tip=<%=Request.QueryString("tip")%>" method=post>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="center"><B>Bilgilerin Yerleþtirileceði Database :</B> <%=veritabani_adi%>.mdb</td>
	</tr>
	<tr>
		<td align="center"><B>Bilgilerin Alýnacaðý Database :<BR>DB/<input type="text" name="eski_mdb" class="form1" size="40">.mdb</B></td>
	</tr>
</tr>
<%		If Request.QueryString("tip") = "mininuke" Then%>
<SCRIPT language=JavaScript src="inc/mn2.js" 
type=text/javascript></SCRIPT>
	<tr>
		<td colspan="2" align="center">Mini Nuke Script sisteminden Maxinuke V <%=v%> sistemine çekilmesini istediðiniz veriler.</td>
	</tr>
	<tr>
		<td  colspan="2" align="center"><INPUT class="submit" onclick=tumusec(document.peno) type=button value="Hepsini Sec" name=on>&nbsp;<INPUT class="submit" onclick=tumuiptal(document.peno) type=button value="Hepsini Sil" name=off>&nbsp;</td>
	</tr>
	<tr>
		<td align="center">
			<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
			<TR>
				<TD>
				<INPUT TYPE="checkbox" NAME="makale" value="evet"> Makaleler<BR>
				<INPUT TYPE="checkbox" NAME="avatar" value="evet"> Avatarlar *<BR>
				<INPUT TYPE="checkbox" NAME="yasakli" value="evet"> Banlý Yasaklý Kiþiler<BR>
				<INPUT TYPE="checkbox" NAME="reklam" value="evet"> Banner Reklamlar *<BR>
				<INPUT TYPE="checkbox" NAME="blok" value="evet"> Bloklar<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="yorum" value="evet"> Yorumlar<BR>
				<INPUT TYPE="checkbox" NAME="dosya" value="evet"> Dosyalar<BR>
				<INPUT TYPE="checkbox" NAME="resim" value="evet"> Resimler<BR>
				<INPUT TYPE="checkbox" NAME="arkadas" value="evet"> Arkadaþlar<BR>
				<INPUT TYPE="checkbox" NAME="ziyaret" value="evet"> Ziyaretçi Görüþleri<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="menu" value="evet"> Menü Kategorileri<BR>
				<INPUT TYPE="checkbox" NAME="mesaj" value="evet"> Mesajlar<BR>
				<INPUT TYPE="checkbox" NAME="modul" value="evet"> Moduller<BR>
				<INPUT TYPE="checkbox" NAME="haber" value="evet"> Haberler<BR>
				<INPUT TYPE="checkbox" NAME="duyuru" value="evet"> Duyurular<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="sayfa" value="evet"> Sayfalar<BR>
				<INPUT TYPE="checkbox" NAME="anket" value="evet"> Anketler<BR>
				<INPUT TYPE="checkbox" NAME="ayar" value="evet"> Ayarlar<BR>
				<INPUT TYPE="checkbox" NAME="tema" value="evet"> Temalar<BR>
				<INPUT TYPE="checkbox" NAME="istatistik" value="evet"> Ýstatistikler
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="FIXED" value="evet"> Sabit Haberler<BR>
				<INPUT TYPE="checkbox" NAME="uyeler" value="evet"> Uyeler<BR>
				<INPUT TYPE="checkbox" NAME="MENU_CATS" value="evet"> Menu Kategorileri<BR>
				<INPUT TYPE="checkbox" NAME="MENU_LINKS" value="evet"> Menu Linkleri<BR>
				</TD>
			</TR>
			</TABLE>
		</td>
	</tr>
	<tr>
		<td  colspan="2" align="center"><INPUT class="submit" onclick=tumusec(document.peno) type=button value="Hepsini Sec" name=on>&nbsp;<INPUT class="submit" onclick=tumuiptal(document.peno) type=button value="Hepsini Sil" name=off>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><INPUT TYPE="submit" style="width=100%" CLASS="submit" value="VERILERI AKTAR"></td>
	</tr>
</table>
</form>
<%
		End if
		If Request.QueryString("tip") = "maxinuke" Then%>
<SCRIPT language=JavaScript src="inc/mn2.js" 
type=text/javascript></SCRIPT>
	<tr>
		<td colspan="2" align="center">MAXINUKE Eski Script sisteminden Maxinuke V <%=v%> sistemine çekilmesini istediðiniz veriler.</td>
	</tr>
	<tr>
		<td  colspan="2" align="center"><INPUT class="submit" onclick=tumusec(document.peno) type=button value="Hepsini Sec" name=on>&nbsp;<INPUT class="submit" onclick=tumuiptal(document.peno) type=button value="Hepsini Sil" name=off>&nbsp;</td>
	</tr>
	<tr>
		<td align="center">
			<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
			<TR>
				<TD>
				<INPUT TYPE="checkbox" NAME="makale" value="evet"> Makaleler<BR>
				<INPUT TYPE="checkbox" NAME="avatar" value="evet"> Avatarlar<BR>
				<INPUT TYPE="checkbox" NAME="yasakli" value="evet"> Banlý Yasaklý Kiþiler<BR>
				<INPUT TYPE="checkbox" NAME="reklam" value="evet"> Banner Reklamlar<BR>
				<INPUT TYPE="checkbox" NAME="blok" value="evet"> Bloklar<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="yorum" value="evet"> Yorumlar<BR>
				<INPUT TYPE="checkbox" NAME="dosya" value="evet"> Dosyalar<BR>
				<INPUT TYPE="checkbox" NAME="resim" value="evet"> Resimler<BR>
				<INPUT TYPE="checkbox" NAME="arkadas" value="evet"> Arkadaþlar<BR>
				<INPUT TYPE="checkbox" NAME="ziyaret" value="evet"> Ziyaretçi Görüþleri<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="menu" value="evet"> Menü Kategorileri<BR>
				<INPUT TYPE="checkbox" NAME="mesaj" value="evet"> Mesajlar<BR>
				<INPUT TYPE="checkbox" NAME="modul" value="evet"> Moduller<BR>
				<INPUT TYPE="checkbox" NAME="haber" value="evet"> Haberler<BR>
				<INPUT TYPE="checkbox" NAME="duyuru" value="evet"> Duyurular<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="sayfa" value="evet"> Sayfalar<BR>
				<INPUT TYPE="checkbox" NAME="anket" value="evet"> Anketler<BR>
				<INPUT TYPE="checkbox" NAME="ayar" value="evet"> Ayarlar<BR>
				<INPUT TYPE="checkbox" NAME="tema" value="evet"> Temalar<BR>
				<INPUT TYPE="checkbox" NAME="istatistik" value="evet"> Ýstatistikler
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="FIXED" value="evet"> Sabit Haberler<BR>
				<INPUT TYPE="checkbox" NAME="uyeler" value="evet"> Uyeler<BR>
				<INPUT TYPE="checkbox" NAME="MENU_CATS" value="evet"> Menu Kategorileri<BR>
				<INPUT TYPE="checkbox" NAME="MENU_LINKS" value="evet"> Menu Linkleri<BR>
				<INPUT TYPE="checkbox" NAME="ekart" value="evet"> E-Kartlar<BR>
				</TD>
				<TD>
				<INPUT TYPE="checkbox" NAME="forum" value="evet"> Forumlar<BR>
				<INPUT TYPE="checkbox" NAME="fikra" value="evet"> Fýkralar<BR>
				<INPUT TYPE="checkbox" NAME="video" value="evet"> Videolar<BR>
				</TD>
			</TR>
			</TABLE>
		</td>
	</tr>
	<tr>
		<td  colspan="2" align="center"><INPUT class="submit" onclick=tumusec(document.peno) type=button value="Hepsini Sec" name=on>&nbsp;<INPUT class="submit" onclick=tumuiptal(document.peno) type=button value="Hepsini Sil" name=off>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><INPUT TYPE="submit" style="width=100%" CLASS="submit" value="VERILERI AKTAR"></td>
	</tr>
</table>
</form>
<%
		End if
	elseIf Request.QueryString("eylem") = "basla" Then
	eskisiadi=Request.Form("eski_mdb")
	Set mdbkontrol = Server.CreateObject("Scripting.FileSystemObject")
		IF Not mdbkontrol.FileExists(""&Server.Mappath("db/"&eskisiadi&".mdb")&"") = True Then
		response.write "<div class=td_menu><BR><BR><BR><CENTER>DB klasoru icerisinde Belirttiginiz adda veritabani bulunamadi<BR><BR>Verlerin alinacagi veritabaniyla verilerin yerlestirilecegi veritabaninin ayni klasorde oldugundan veya eski veritabani adini dogru yazdiginizdan emin olun.<BR><BR><BR><input type=button value=GERI GIT onclick=javascript:history.back(); class=submit></CENTER></div>"
		response.end
		End if
	Set mdbkontrol = Nothing
	Set eskibaglanti = Server.CreateObject("ADODB.Connection")
	eskibaglanti.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/"&eskisiadi&".mdb")
		If Request.QueryString("tip") = "mininuke" Then
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="center">
<%
Session("makale")=Request.form("makale")
Session("avatar")=Request.form("avatar")
Session("yasakli")=Request.form("yasakli")
Session("reklam")=Request.form("reklam")
Session("blok")=Request.form("blok")
Session("yorum")=Request.form("yorum")
Session("dosya")=Request.form("dosya")
Session("resim")=Request.form("resim")
Session("arkadas")=Request.form("arkadas")
Session("ziyaret")=Request.form("ziyaret")
Session("menu")=Request.form("menu")
Session("mesaj")=Request.form("mesaj")
Session("modul")=Request.form("modul")
Session("haber")=Request.form("haber")
Session("duyuru")=Request.form("duyuru")
Session("sayfa")=Request.form("sayfa")
Session("anket")=Request.form("anket")
Session("ayar")=Request.form("ayar")
Session("tema")=Request.form("tema")
Session("istatistik")=Request.form("istatistik")
Session("FIXED")=Request.form("FIXED")
Session("uyeler")=Request.form("uyeler")
Session("MENU_CATS")=Request.form("MENU_CATS")
Session("MENU_LINKS")=Request.form("MENU_LINKS")
			If Session("makale") = "" And Session("avatar") = "" And Session("yasakli") = "" And Session("reklam") = "" And Session("blok") = "" And Session("yorum") = "" And Session("dosya") = "" And Session("resim") = "" And Session("arkadas") = "" And Session("ziyaret") = "" And Session("menu") = "" And Session("mesaj") = "" And Session("modul") = "" And Session("haber") = "" And Session("duyuru") = "" And Session("sayfa") = "" And Session("anket") = "" And Session("ayar") = "" And Session("tema") = "" And Session("istatistik") = "" And Session("FIXED") = "" And Session("uyeler") = "" And Session("MENU_CATS") = "" And Session("MENU_LINKS") = "" then
			response.write "<BR><BR><BR>Aktarým iþlemini yapabilmek için en az 1 adet alan seçmelisiniz.<BR><BR><BR><input type=button value=GERI GIT onclick=javascript:history.back(); class=submit>"
			Else
			response.write "<B>Mini Nuke Script sisteminden Maxinuke V "&v&" sistemine veriler çekilmeye baþlandý<BR>Lütfen iþlem tamamlandý uyarýsýný alana kadar bu sayfayý kapatmayýn.<BR>Ýþlem süresi; aktarýlacak verilerinizin çokluðuna ve internet hýzýnýza göre deðiþkenlik gösterebilir.</B><IMG SRC=images/yukleme.gif WIDTH=100% HEIGHT=10>"
				If Session("makale") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM ARTICLE_CATS",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM ARTICLE_CATS",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name") 
				yerlestir_kat("cat_info") = "Kategori Açýklamasý Belirtilmemiþ"
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Makale Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM ARTICLES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM ARTICLES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("a_title") = aliver("a_title") 
				yerlestir("a_article") = aliver("a_article")
				yerlestir("a_date") = aliver("a_date")
				yerlestir("a_writer") = aliver("a_writer")
				yerlestir("cat_id") = aliver("cat_id")
				yerlestir("hit") = aliver("hit")
				yerlestir("point") = aliver("point")
				yerlestir("a_approved") = aliver("a_approved")
				yerlestir("p_img") = ""
				yerlestir("anahtar") = "-"
				yerlestir("align") = "left"
				yerlestir("ekleyen") = "Admin"
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Makaleler Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_kat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("avatar") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM AVATARS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM AVATARS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("a_name") = aliver("a_name") 
				yerlestir("a_img") = aliver("a_img")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Avatarlar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("yasakli") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM BANNED_IPS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM BANNED_IPS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("b_ip") = aliver("b_ip") 
				yerlestir("b_date") = aliver("b_date")
				yerlestir("sebep") = "Belirtilmemiþ"
				yerlestir("banlayan") = "Belirtilmemiþ"
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Banlanmis - Yasaklanmis Kisiler Listesi Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("reklam") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM BANNERS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM BANNERS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("ban_url") = aliver("ban_url") 
				yerlestir("go_url") = aliver("go_url")
				yerlestir("hit") = aliver("hit")
				yerlestir("active") = aliver("active")
				yerlestir("active") = aliver("active")
				yerlestir("position") = aliver("position")
				yerlestir("limit") = aliver("limit")
				yerlestir("show") = aliver("show")
				yerlestir("password") = aliver("password")
				If aliver("ban_type") = "Normal" Then
				ban_type = "1"
				elseIf aliver("ban_type") = "Flash" Then
				ban_type = "2"
				elseIf aliver("ban_type") = "Kod" Then
				ban_type = "3"
				End if
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Reklamlariniz Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("blok") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM BLOCKS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM bloklar",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("b_name") = aliver("b_name") 
				yerlestir("b_content") = aliver("b_content")
				yerlestir("b_align") = aliver("b_align")
				yerlestir("b_include") = aliver("b_include")
				yerlestir("b_active") = aliver("b_active")
				yerlestir("b_queue") = aliver("b_queue")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Bloklariniz Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("yorum") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM COMMENTS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM COMMENTS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("writer") = aliver("writer") 
				yerlestir("comment") = aliver("comment")
				yerlestir("cdate") = aliver("cdate")
				yerlestir("nid") = aliver("nid")
				yerlestir("page") = aliver("page")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Yorumlariniz Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("dosya") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM DOWN_CATS",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM DOWN_CATS",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("c_name") = aliver_kat("c_name") 
				yerlestir_kat("cat_info") = "Belirtilmemis"
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Dosya Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM DOWNLOADS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM DOWNLOADS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("p_name") = aliver("p_name") 
				yerlestir("p_info") = aliver("p_info")
				yerlestir("p_size") = aliver("p_size")
				yerlestir("p_hit") = aliver("p_hit")
				yerlestir("p_url") = aliver("p_url")
				yerlestir("cid") = aliver("cid")
				yerlestir("p_date") = aliver("p_date")
				yerlestir("p_author") = aliver("p_author")
				yerlestir("point") = aliver("point")
				yerlestir("p_img") = aliver("p_img")
				yerlestir("anahtar") = "-"
				yerlestir("onay") = aliver("p_approved")
				yerlestir("align") = "Left"
				yerlestir("ekleyen") = "Admin"
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Dosyalar Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("FIXED") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM FIXED",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM FIXED",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("sf_baslik") = aliver("sf_baslik") 
				yerlestir("f_metin") = aliver("f_metin")
				yerlestir("resim") = aliver("resim")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Sabit Haberler Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("resim") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM kategori",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM rg_cat",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("kategori") 
				yerlestir_kat("cat_info") = aliver_kat("statement") 
				yerlestir_kat("cat_image") = ""
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Resim Kategorileri Aktarýldý !</div>"				
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM fotograf",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM rg",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("isim") 
				yerlestir("metin") = aliver("aciklama")
				yerlestir("ekleyen") = "Admin"
				yerlestir("tarih") = aliver("tarih")
				yerlestir("onay") = true
				yerlestir("resim") = aliver("buyuk")
				yerlestir("puan") = "0"
				yerlestir("katilimci") = "0"
				yerlestir("bakilma") = aliver("hit")
				yerlestir("cat") = aliver("kategori")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Resimler Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("arkadas") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM FRIENDS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM FRIENDS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("MEMBER") = aliver("MEMBER") 
				yerlestir("FRIEND") = aliver("FRIEND")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Arkadaslar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("ziyaret") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM GUESTBOOK",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM GUESTBOOK",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("NAME") = aliver("NAME") 
				yerlestir("EMAIL") = aliver("EMAIL")
				yerlestir("YAZI") = aliver("YAZI")
				yerlestir("tarih") = Now()			
				yerlestir("onay") = true
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Arkadaslar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
'				If Session("menu") = "evet" Then
'				Set aliver = Server.CreateObject("ADODB.RecordSet")
'				aliver.open "Select * FROM GUESTBOOK",eskibaglanti,3,3
'				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
'				yerlestir.open "Select * FROM GUESTBOOK",Connection,3,3
'				Do Until aliver.EoF
'				yerlestir.AddNew
'				yerlestir("NAME") = aliver("NAME") 
'				yerlestir("EMAIL") = aliver("EMAIL")
'				yerlestir("YAZI") = aliver("YAZI")
'				yerlestir("tarih") = Now()			
''				yerlestir("onay") = true
'				yerlestir.Update
'				aliver.MoveNext
'				Loop
'				Response.Write "<div align=center class=td_menu>Arkadaslar Aktarýldý !</div>"
'				yerlestir.Close
'				Set yerlestir = Nothing
'				aliver.Close
'				Set aliver = Nothing
'				End If
				If Session("uyeler") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MEMBERS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MEMBERS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("kul_adi") = aliver("kul_adi") 
				yerlestir("sifre") = aliver("sifre")
				yerlestir("email") = aliver("email")
				yerlestir("isim") = aliver("isim")
				yerlestir("g_soru") = aliver("g_soru")
				yerlestir("g_cevap") = aliver("g_cevap")
				yerlestir("sehir") = aliver("sehir")
				yerlestir("meslek") = aliver("meslek")
				yerlestir("cinsiyet") = aliver("cinsiyet")
				yerlestir("yas") = aliver("yas")
				yerlestir("imza") = aliver("imza")
				yerlestir("login_sayisi") = aliver("login_sayisi")
				yerlestir("uyelik_tarihi") = aliver("uyelik_tarihi")
				yerlestir("son_tarih") = aliver("son_tarih")
				yerlestir("seviye") = aliver("seviye")
				yerlestir("msg_sayisi") = aliver("msg_sayisi")
				yerlestir("last_ip") = aliver("last_ip")
				yerlestir("u_theme") = "b"
				yerlestir("u_avatar") = aliver("u_avatar")
				yerlestir("u_online") = false
				yerlestir("u_browser") = aliver("u_browser")
				yerlestir("u_ws") = aliver("u_ws")
				yerlestir("onay") = false
				yerlestir("gk") = "123456789"
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Uyeler Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("MENU_CATS") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MENU_CATS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MENU_CATS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("mc_title") = aliver("mc_title") 
				yerlestir("mc_queue") = aliver("mc_queue")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Menu Kategorileri Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("MENU_LINKS") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MENU_LINKS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MENU_LINKS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("m_name") = aliver("m_name")
				yerlestir("m_url") = aliver("m_url") 
				yerlestir("m_status") = aliver("m_status")
				yerlestir("m_style") = aliver("m_style")
				yerlestir("m_cat") = aliver("m_cat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Menu Linkleri Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("mesaj") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MESSAGES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MESSAGES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("from") = aliver("from")
				yerlestir("to") = aliver("to") 
				yerlestir("msg") = aliver("msg")
				yerlestir("mdate") = aliver("mdate")
				yerlestir("read") = aliver("read")
				yerlestir("subject") = aliver("subject")
				yerlestir("type") = aliver("type")
				yerlestir("see") = aliver("see")
				yerlestir("msave") = aliver("msave")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Uyeler Arasi Ic Mesajlar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("modul") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MODULES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM moduller",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("mname") = aliver("mname")
				yerlestir("mdir") = aliver("mdir") 
				yerlestir("mactive") = aliver("mactive")
				yerlestir("mems") = aliver("mems")
				yerlestir("mcat") = aliver("mcat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Moduller Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("haber") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM NEWS_CATS",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM NEWS_CATS",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name")
				yerlestir_kat("cat_info") = aliver_kat("cat_info")
				yerlestir_kat("cat_image") = aliver_kat("cat_image")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Haber Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM NEWS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM NEWS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("baslik")
				yerlestir("metin") = aliver("metin") 
				yerlestir("ekleyen") = aliver("ekleyen")
				yerlestir("tarih") = aliver("tarih")
				yerlestir("onay") = aliver("onay")
				yerlestir("resim") = aliver("resim")
				yerlestir("puan") = aliver("puan")
				yerlestir("katilimci") = aliver("katilimci")
				yerlestir("okunma") = aliver("okunma")
				yerlestir("cat") = aliver("cat")
				yerlestir("align") = "left"
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Haberler Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_kat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("duyuru") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM NOTICES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM NOTICES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("n_text") = aliver("n_text")
				yerlestir("n_date") = aliver("n_date")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Duyurular Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("sayfa") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM PAGES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM PAGES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("p_title") = aliver("p_title")
				yerlestir("p_cat") = aliver("p_cat")
				yerlestir("resim") = ""
				yerlestir("p_content") = aliver("p_content")
				yerlestir("onay") = true
				yerlestir("align") = "left"
				yerlestir("mems") = "Admin"
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Sayfalar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("anket") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM POLL_ALTERNATIVES",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM POLL_ALTERNATIVES",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("pid") = aliver_kat("pid")
				yerlestir_kat("alternative") = aliver_kat("alternative")
				yerlestir_kat("a_counter") = aliver_kat("a_counter")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Anket Yanitlari Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM POLLS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM POLLS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("p_question") = aliver("p_question") 
				yerlestir("p_date") = aliver("p_date")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Anket Sorulari Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("ayar") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM SETTINGS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM SETTINGS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("site_isim") = aliver("site_isim") 
				yerlestir("site_adres") = aliver("site_adres")
				yerlestir("site_logo") = aliver("site_logo")
				yerlestir("site_slogan") = aliver("site_slogan")
				yerlestir("flood_time") = aliver("flood_time")
				yerlestir("site_mail") = aliver("site_mail")
				yerlestir("prg_sayisi") = aliver("prg_sayisi")
				yerlestir("msg_limit") = aliver("msg_limit")
				yerlestir("theme") = "b"
				yerlestir("np_msg") = aliver("np_msg")
				yerlestir("fpass") = aliver("fpass")
				yerlestir("s_lcid") = aliver("s_lcid")
				yerlestir("google") = true
				yerlestir("onaylama") = false
				yerlestir("logo_tipi") = false
				yerlestir("son_dakika") = true
				yerlestir("f_posts") = "20"
				yerlestir("f_replies") = "20"
				yerlestir("Keywords") = "-"
				yerlestir("sd") = true
				yerlestir("sm") = true
				yerlestir("su") = true
				yerlestir("sf") = true
				yerlestir("se") = true
				yerlestir("sfb") = true
				yerlestir("sr") = true
				yerlestir("so") = true
				yerlestir("sv") = true
				yerlestir("rd") = true
				yerlestir("rm") = true
				yerlestir("ru") = true
				yerlestir("rf") = true
				yerlestir("re") = true
				yerlestir("rr") = true
				yerlestir("ro") = true
				yerlestir("rv") = true
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Ayarlar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("istatistik") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM WEBCOUNTER",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM WEBCOUNTER",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("today") = aliver("today") 
				yerlestir("cdate") = aliver("cdate")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Istatistikler Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
%>
<meta http-equiv="refresh" content="5;url=?eylem=ok">
Ýþlem sonuçlandýrýlýyor . . .
<%
			End if
		End if
		If Request.QueryString("tip") = "maxinuke" Then
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="center">
<%
Session("makale")=Request.form("makale")
Session("avatar")=Request.form("avatar")
Session("yasakli")=Request.form("yasakli")
Session("reklam")=Request.form("reklam")
Session("blok")=Request.form("blok")
Session("yorum")=Request.form("yorum")
Session("dosya")=Request.form("dosya")
Session("resim")=Request.form("resim")
Session("arkadas")=Request.form("arkadas")
Session("ziyaret")=Request.form("ziyaret")
Session("menu")=Request.form("menu")
Session("mesaj")=Request.form("mesaj")
Session("modul")=Request.form("modul")
Session("haber")=Request.form("haber")
Session("duyuru")=Request.form("duyuru")
Session("sayfa")=Request.form("sayfa")
Session("anket")=Request.form("anket")
Session("ayar")=Request.form("ayar")
Session("tema")=Request.form("tema")
Session("istatistik")=Request.form("istatistik")
Session("FIXED")=Request.form("FIXED")
Session("uyeler")=Request.form("uyeler")
Session("MENU_CATS")=Request.form("MENU_CATS")
Session("MENU_LINKS")=Request.form("MENU_LINKS")
Session("ekart")=Request.form("ekart")
Session("forum")=Request.form("forum")
Session("fikra")=Request.form("fikra")
Session("video")=Request.form("video")

			If Session("makale") = "" And Session("avatar") = "" And Session("yasakli") = "" And Session("reklam") = "" And Session("blok") = "" And Session("yorum") = "" And Session("dosya") = "" And Session("resim") = "" And Session("arkadas") = "" And Session("ziyaret") = "" And Session("menu") = "" And Session("mesaj") = "" And Session("modul") = "" And Session("haber") = "" And Session("duyuru") = "" And Session("sayfa") = "" And Session("anket") = "" And Session("ayar") = "" And Session("tema") = "" And Session("istatistik") = "" And Session("FIXED") = "" And Session("uyeler") = "" And Session("MENU_CATS") = "" And Session("MENU_LINKS") = "" And Session("ekart") = "" And Session("forum") = "" And Session("fikra") = "" And Session("video") = "" then
			response.write "<BR><BR><BR>Aktarým iþlemini yapabilmek için en az 1 adet alan seçmelisiniz.<BR><BR><BR><input type=button value=GERI GIT onclick=javascript:history.back(); class=submit>"
			Else
			response.write "<B>Mini Nuke Script sisteminden Maxinuke V "&v&" sistemine veriler çekilmeye baþlandý<BR>Lütfen iþlem tamamlandý uyarýsýný alana kadar bu sayfayý kapatmayýn.<BR>Ýþlem süresi; aktarýlacak verilerinizin çokluðuna ve internet hýzýnýza göre deðiþkenlik gösterebilir.</B><IMG SRC=images/yukleme.gif WIDTH=100% HEIGHT=10>"
				If Session("makale") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM ARTICLE_CATS",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM ARTICLE_CATS",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name") 
				yerlestir_kat("cat_info") = aliver_kat("cat_info") 
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Makale Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM ARTICLES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM ARTICLES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("a_title") = aliver("a_title") 
				yerlestir("a_article") = aliver("a_article")
				yerlestir("a_date") = aliver("a_date")
				yerlestir("a_writer") = aliver("a_writer")
				yerlestir("cat_id") = aliver("cat_id")
				yerlestir("hit") = aliver("hit")
				yerlestir("point") = aliver("point")
				yerlestir("anahtar") = aliver("anahtar")
				yerlestir("a_approved") = aliver("a_approved")
				yerlestir("ekleyen") = aliver("ekleyen")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Makaleler Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_kat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("avatar") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM AVATARS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM AVATARS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("a_name") = aliver("a_name") 
				yerlestir("a_img") = aliver("a_img")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Avatarlar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("yasakli") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM BANNED_IPS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM BANNED_IPS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("b_ip") = aliver("b_ip") 
				yerlestir("b_date") = aliver("b_date")
				yerlestir("sebep") = aliver("sebep")
				yerlestir("banlayan") = aliver("banlayan")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Banlanmis - Yasaklanmis Kisiler Listesi Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("reklam") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM BANNERS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM BANNERS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("ban_url") = aliver("ban_url") 
				yerlestir("go_url") = aliver("go_url")
				yerlestir("hit") = aliver("hit")
				yerlestir("active") = aliver("active")
				yerlestir("position") = aliver("position")
				yerlestir("limit") = aliver("limit")
				yerlestir("show") = aliver("show")
				yerlestir("password") = aliver("password")
				yerlestir("ban_type") = aliver("ban_type")
				yerlestir("baslangic") = aliver("baslangic")
				yerlestir("bitis") = aliver("bitis")
				yerlestir("gunlimit") = aliver("gunlimit")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Reklamlariniz Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("blok") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM bloklar",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM bloklar",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("b_name") = aliver("b_name") 
				yerlestir("b_content") = aliver("b_content")
				yerlestir("b_align") = aliver("b_align")
				yerlestir("b_include") = aliver("b_include")
				yerlestir("b_active") = aliver("b_active")
				yerlestir("b_queue") = aliver("b_queue")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Bloklariniz Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("yorum") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM COMMENTS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM COMMENTS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("writer") = aliver("writer") 
				yerlestir("comment") = aliver("comment")
				yerlestir("cdate") = aliver("cdate")
				yerlestir("nid") = aliver("nid")
				yerlestir("page") = aliver("page")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Yorumlariniz Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("dosya") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM DOWN_CATS",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM DOWN_CATS",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("c_name") = aliver_kat("c_name") 
				yerlestir_kat("cat_info") = aliver_kat("cat_info")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Dosya Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM DOWNLOADS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM DOWNLOADS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("p_name") = aliver("p_name") 
				yerlestir("p_info") = aliver("p_info")
				yerlestir("p_size") = aliver("p_size")
				yerlestir("p_hit") = aliver("p_hit")
				yerlestir("p_url") = aliver("p_url")
				yerlestir("cid") = aliver("cid")
				yerlestir("p_date") = aliver("p_date")
				yerlestir("p_author") = aliver("p_author")
				yerlestir("point") = aliver("point")
				yerlestir("anahtar") = aliver("anahtar")
				yerlestir("onay") = aliver("onay")
				yerlestir("ekleyen") = aliver("ekleyen")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Dosyalar Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("FIXED") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM FIXED",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM FIXED",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("sf_baslik") = aliver("sf_baslik") 
				yerlestir("f_metin") = aliver("f_metin")
				yerlestir("resim") = aliver("resim")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Sabit Haberler Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("resim") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM rg_cat",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM rg_cat",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name") 
				yerlestir_kat("cat_info") = aliver_kat("cat_info") 
				yerlestir_kat("cat_image") = aliver_kat("cat_image") 
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Resim Kategorileri Aktarýldý !</div>"				
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM rg",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM rg",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("baslik") 
				yerlestir("metin") = aliver("metin")
				yerlestir("ekleyen") = aliver("ekleyen")
				yerlestir("tarih") = aliver("tarih")
				yerlestir("onay") = aliver("onay")
				yerlestir("puan") = aliver("puan")
				yerlestir("katilimci") = aliver("katilimci")
				yerlestir("bakilma") = aliver("bakilma")
				yerlestir("cat") = aliver("cat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Resimler Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("arkadas") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM FRIENDS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM FRIENDS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("MEMBER") = aliver("MEMBER") 
				yerlestir("FRIEND") = aliver("FRIEND")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Arkadaslar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("ziyaret") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM GUESTBOOK",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM GUESTBOOK",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("NAME") = aliver("NAME") 
				yerlestir("EMAIL") = aliver("EMAIL")
				yerlestir("YAZI") = aliver("YAZI")
				yerlestir("tarih") = aliver("tarih")
				yerlestir("onay") = aliver("onay")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Arkadaslar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
'				If Session("menu") = "evet" Then
'				Set aliver = Server.CreateObject("ADODB.RecordSet")
'				aliver.open "Select * FROM GUESTBOOK",eskibaglanti,3,3
'				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
'				yerlestir.open "Select * FROM GUESTBOOK",Connection,3,3
'				Do Until aliver.EoF
'				yerlestir.AddNew
'				yerlestir("NAME") = aliver("NAME") 
'				yerlestir("EMAIL") = aliver("EMAIL")
'				yerlestir("YAZI") = aliver("YAZI")
'				yerlestir("tarih") = Now()			
'				yerlestir("onay") = true
'				yerlestir.Update
'				aliver.MoveNext
'				Loop
''				Response.Write "<div align=center class=td_menu>Arkadaslar Aktarýldý !</div>"
'				yerlestir.Close
'				Set yerlestir = Nothing
'				aliver.Close
'				Set aliver = Nothing
'				End If
				If Session("uyeler") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MEMBERS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MEMBERS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("kul_adi") = aliver("kul_adi") 
				yerlestir("sifre") = aliver("sifre")
				yerlestir("email") = aliver("email")
				yerlestir("isim") = aliver("isim")
				yerlestir("g_soru") = aliver("g_soru")
				yerlestir("g_cevap") = aliver("g_cevap")
				yerlestir("sehir") = aliver("sehir")
				yerlestir("meslek") = aliver("meslek")
				yerlestir("cinsiyet") = aliver("cinsiyet")
				yerlestir("yas") = aliver("yas")
				yerlestir("imza") = aliver("imza")
				yerlestir("login_sayisi") = aliver("login_sayisi")
				yerlestir("uyelik_tarihi") = aliver("uyelik_tarihi")
				yerlestir("son_tarih") = aliver("son_tarih")
				yerlestir("seviye") = aliver("seviye")
				yerlestir("msg_sayisi") = aliver("msg_sayisi")
				yerlestir("last_ip") = aliver("last_ip")
				yerlestir("u_theme") = aliver("u_theme")
				yerlestir("u_avatar") = aliver("u_avatar")
				yerlestir("u_online") = aliver("u_online")
				yerlestir("u_browser") = aliver("u_browser")
				yerlestir("u_ws") = aliver("u_ws")
				yerlestir("onay") = aliver("onay")
				yerlestir("gk") = aliver("gk")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Uyeler Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("MENU_CATS") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MENU_CATS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MENU_CATS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("mc_title") = aliver("mc_title") 
				yerlestir("mc_queue") = aliver("mc_queue")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Menu Kategorileri Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("MENU_LINKS") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MENU_LINKS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MENU_LINKS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("m_name") = aliver("m_name")
				yerlestir("m_url") = aliver("m_url") 
				yerlestir("m_status") = aliver("m_status")
				yerlestir("m_style") = aliver("m_style")
				yerlestir("m_cat") = aliver("m_cat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Menu Linkleri Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("mesaj") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM MESSAGES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM MESSAGES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("from") = aliver("from")
				yerlestir("to") = aliver("to") 
				yerlestir("msg") = aliver("msg")
				yerlestir("mdate") = aliver("mdate")
				yerlestir("read") = aliver("read")
				yerlestir("subject") = aliver("subject")
				yerlestir("type") = aliver("type")
				yerlestir("see") = aliver("see")
				yerlestir("msave") = aliver("msave")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Uyeler Arasi Ic Mesajlar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("modul") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM moduller",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM moduller",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("mname") = aliver("mname")
				yerlestir("mdir") = aliver("mdir") 
				yerlestir("mactive") = aliver("mactive")
				yerlestir("mems") = aliver("mems")
				yerlestir("mcat") = aliver("mcat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Moduller Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("haber") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM NEWS_CATS",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM NEWS_CATS",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name")
				yerlestir_kat("cat_info") = aliver_kat("cat_info")
				yerlestir_kat("cat_image") = aliver_kat("cat_image")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Haber Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM NEWS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM NEWS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("baslik")
				yerlestir("metin") = aliver("metin") 
				yerlestir("ekleyen") = aliver("ekleyen")
				yerlestir("tarih") = aliver("tarih")
				yerlestir("onay") = aliver("onay")
				yerlestir("resim") = aliver("resim")
				yerlestir("puan") = aliver("puan")
				yerlestir("katilimci") = aliver("katilimci")
				yerlestir("okunma") = aliver("okunma")
				yerlestir("cat") = aliver("cat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Haberler Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_kat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("duyuru") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM NOTICES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM NOTICES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("n_text") = aliver("n_text")
				yerlestir("n_date") = aliver("n_date")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Duyurular Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("sayfa") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM PAGES",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM PAGES",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("p_title") = aliver("p_title")
				yerlestir("p_cat") = aliver("p_cat")
				yerlestir("p_content") = aliver("p_content")
				yerlestir("onay") = aliver("onay")
				yerlestir("mems") = aliver("mems")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Sayfalar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("anket") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM POLL_ALTERNATIVES",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM POLL_ALTERNATIVES",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("pid") = aliver_kat("pid")
				yerlestir_kat("alternative") = aliver_kat("alternative")
				yerlestir_kat("a_counter") = aliver_kat("a_counter")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Anket Yanitlari Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM POLLS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM POLLS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("p_question") = aliver("p_question") 
				yerlestir("p_date") = aliver("p_date")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Anket Sorulari Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("ayar") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM SETTINGS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM SETTINGS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("site_isim") = aliver("site_isim") 
				yerlestir("site_adres") = aliver("site_adres")
				yerlestir("site_logo") = aliver("site_logo")
				yerlestir("site_slogan") = aliver("site_slogan")
				yerlestir("flood_time") = aliver("flood_time")
				yerlestir("site_mail") = aliver("site_mail")
				yerlestir("prg_sayisi") = aliver("prg_sayisi")
				yerlestir("msg_limit") = aliver("msg_limit")
				yerlestir("theme") = aliver("theme")
				yerlestir("np_msg") = aliver("np_msg")
				yerlestir("fpass") = aliver("fpass")
				yerlestir("s_lcid") = aliver("s_lcid")
				yerlestir("google") = aliver("google")
				yerlestir("onaylama") = aliver("onaylama")
				yerlestir("logo_tipi") = aliver("logo_tipi")
				yerlestir("son_dakika") = aliver("son_dakika")
				yerlestir("f_posts") = aliver("f_posts")
				yerlestir("f_replies") = aliver("f_replies")
				yerlestir("Keywords") = aliver("Keywords")
				yerlestir("sd") = aliver("sd")
				yerlestir("sm") = aliver("sm")
				yerlestir("su") = aliver("su")
				yerlestir("sf") = aliver("sf")
				yerlestir("se") = aliver("se")
				yerlestir("sfb") = aliver("sfb")
				yerlestir("sr") = aliver("sr")
				yerlestir("so") = aliver("so")
				yerlestir("sv") = aliver("sv")
				yerlestir("rd") = aliver("rd")
				yerlestir("rm") = aliver("rm")
				yerlestir("ru") = aliver("ru")
				yerlestir("rf") = aliver("rf")
				yerlestir("re") = aliver("re")
				yerlestir("rr") = aliver("rr")
				yerlestir("ro") = aliver("ro")
				yerlestir("rv") = aliver("rv")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Ayarlar Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("istatistik") = "evet" Then
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM WEBCOUNTER",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM WEBCOUNTER",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("today") = aliver("today") 
				yerlestir("cdate") = aliver("cdate")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Istatistikler Aktarýldý !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("ekart") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM ekart_kat",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM ekart_kat",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("isim") = aliver_kat("isim") 
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Ekart Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM ekartlar",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM ekartlar",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("katid") = aliver("katid") 
				yerlestir("buyuk") = aliver("buyuk")
				yerlestir("yol") = aliver("yol")
				yerlestir("tarih") = aliver("tarih")
				yerlestir("onay") = aliver("onay")
				yerlestir("hit") = aliver("hit")
				yerlestir("tip") = aliver("tip")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Ekartlar Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM ekart_gonderilen",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM ekart_gonderilen",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("resid") = aliver("resid") 
				yerlestir("kime") = aliver("kime")
				yerlestir("kimden") = aliver("kimden")
				yerlestir("kimdene") = aliver("kimdene")
				yerlestir("kimee") = aliver("kimee")
				yerlestir("baslik") = aliver("baslik")
				yerlestir("mesaj") = aliver("mesaj")
				yerlestir("backcolor") = aliver("backcolor")
				yerlestir("bordercolor") = aliver("bordercolor")
				yerlestir("backid") = aliver("backid")
				yerlestir("midid") = aliver("midid")
				yerlestir("font") = aliver("font")
				yerlestir("haberver") = aliver("haberver")
				yerlestir("okunma") = aliver("okunma")
				yerlestir("createdtime") = aliver("createdtime")
				yerlestir("kartid") = aliver("kartid")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Gonderýlmýs Ekartlar Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				End if
				If Session("forum") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM f_anakategori",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM f_anakategori",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("anakatadi") = aliver_kat("anakatadi") 
				yerlestir_kat("sira") = aliver_kat("sira") 
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Forum Ana Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM f_kategoriler",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM f_kategoriler",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("katad") = aliver("katad") 
				yerlestir("kataciklama") = aliver("kataciklama")
				yerlestir("kilit") = aliver("kilit")
				yerlestir("noentry") = aliver("noentry")
				yerlestir("anakatid") = aliver("anakatid")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Forum Kategorýlerý Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM f_mesajlar",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM f_mesajlar",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("baslik") 
				yerlestir("mmsg") = aliver("mmsg")
				yerlestir("yazan") = aliver("yazan")
				yerlestir("tarih") = aliver("tarih")
				yerlestir("okunma") = aliver("okunma")
				yerlestir("kategoriid") = aliver("kategoriid")
				yerlestir("kilit") = aliver("kilit")
				yerlestir("topic") = aliver("topic")
				yerlestir("html_allow") = aliver("html_allow")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Forum Mesajlarý Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing				
				End if
				If Session("fikra") = "evet" Then
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM fikra_cat",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM fikra_cat",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Fýkra Kategorileri Aktarýldý !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM fikra",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM fikra",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("baslik") 
				yerlestir("metin") = aliver("metin")
				yerlestir("onay") = aliver("onay")
				yerlestir("puan") = aliver("puan")
				yerlestir("katilimci") = aliver("katilimci")
				yerlestir("hit") = aliver("hit")
				yerlestir("cat") = aliver("cat")
				yerlestir("gonderen") = aliver("gonderen")
				yerlestir("eklenme") = aliver("eklenme")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Fýkralar Aktarýldý !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing		
				End if
				If Session("videolar") = "evet" Then
				End if
%>
<meta http-equiv="refresh" content="5;url=?eylem=ok">
Ýþlem sonuçlandýrýlýyor . . .
<%
			End if
		End if
%>
		</TD>
	</TR>
</TABLE>
<%
	eskibaglanti.Close
	Set eskibaglanti = Nothing	
	elseIf Request.QueryString("eylem") = "ok" then
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="center">Belirlemis oldugunuz tum veriler yeni veritabaniniza basariyla aktarildi<BR><BR><B>TEÞEKKÜRLER BAABINDA HAYRINA ALTTAKI REKLAMI TIKLAYAYIM</B><BR><BR>
		<!--#include file="inc/adsense_3.asp" --></td>
	</tr>
</table>
<%
	End if
End If
%>