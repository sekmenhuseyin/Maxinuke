<% Response.buffer = True %>
<!--#include file="inc/database.asp" -->
<!--#include file="inc/functions.asp" -->
<!--#include file="inc/filter.asp" -->
<!--#include file="inc/encryption.asp" -->
<!--#include file="images/temalar/a/info.asp" -->
<%
Set sett = Server.CreateObject("ADODB.RecordSet")
settSQL = "Select * FROM SETTINGS"
sett.open settSQL,Connection,1,3
%>
<html>
<head>
<title>MAXI NUKE V <%=v%> �CRETS�Z KURULUM SAYFASI</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<link rel="stylesheet" type="text/css" href="images/temalar/a/style.css">
<body bgcolor="<%=background%>" leftmargin="5" topmargin="0" marginwidth="0" marginheight="0">
</head>
<%
If Request.QueryString("eylem") = "" Then
	Set anakontrol = Server.CreateObject("Scripting.FileSystemObject")
	IF anakontrol.FileExists(""&Server.Mappath("db/maxinuke.mdb")&"") = True Or anakontrol.FileExists(""&Server.Mappath("../db/maxinuke.mdb")&"") = True Then
%>
<TABLE CLASS="td_menu" width="100%">
	<TR>
		<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
	</TR>
	<TR>
		<TD>Maxi Nuke Version <%=v%> Kurulum Sayfas�na ho�geldiniz,<BR><BR>Kurulum i�lemine ba�lamadan �nce baz� hatalar farkedildi ve a��a��da listelendi, l�tfen gerekli d�zenlemeyi yap�n ve kuruluma tekrar ba�lamay� deneyin. Bu d�zenleme i�lemini yapmazsan�z siteniz sald�rgan ki�ilere kar�� korumas�z olacakt�r.<BR><BR><HR></TD>
	</TR>
	<TR>
		<TD align="center"><CENTER><B><U>BULUNAN HATALAR</U></B></CENTER></TD>
	</TR>
		<%IF anakontrol.FileExists(""&Server.Mappath("db/maxinuke.mdb")&"") = True Then%>
	<TR>
		<TD>
		L�tfen db klas�r� i�erisindeki maxinuke.mdb isimli dosyan�n ad�n� g�venli�iniz i�in de�i�tirin daha sonra inc klas�r� i�indeki database.asp isimli dosyay� not defteri yada benzeri bir text editor programiyla a��n ve 2. sat�rdaki veritabani_adi="<B>maxinuke</B>" yaz�s�ndaki maxinuke yazan k�sm� yeni yapt���n�z d�zenlemeye g�re de�i�tirin.<BR><BR>Daha sonra bu sayfaya gelerek sayfayi yenileyin yada <A HREF="kur.asp"><B>BURAYA</B></A> basin</TD>
	</TR>
		<%
		End If
		IF anakontrol.FileExists(""&Server.Mappath("../db/maxinuke.mdb")&"") = True Then
		%>
	<TR>
		<TD>
		L�tfen bir alt dizindeki korumal� db klas�r� i�erisindeki maxinuke.mdb isimli dosyan�n ad�n� g�venli�iniz i�in de�i�tirin daha sonra inc klas�r� i�indeki database.asp isimli dosyay� not defteri yada benzeri bir text editor programiyla a��n ve 2. sat�rdaki veritabani_adi="<B>maxinuke</B>" yaz�s�ndaki maxinuke yazan k�sm� yeni yapt���n�z d�zenlemeye g�re de�i�tirin.<BR><BR>Daha sonra bu sayfaya gelerek sayfayi yenileyin yada <A HREF="kur.asp"><B>BURAYA</B></A> basin</TD>
	</TR>
		<%End If%>
</TABLE>
	<%
	response.end
	else
	%>
<TABLE CLASS="td_menu" width="100%">
	<TR>
		<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
	</TR>
	<TR>
		<TD align="center">
		<iframe src="http://www.maxinuke.com/guncelleme.asp?v=<%=v%>&t=a" width="100%" height="48" frameborder="0" framespacing="0" scrolling="no" name="maxinuke"></iframe></TD>
	</TR>
	<TR>
		<TD align="center">Maxi Nuke Version <%=v%> Kurulum Sayfas�na ho�geldiniz,<BR><BR>kurulum i�lemine ba�lamadan �nce hatas�z bir kurulum yapabilmek i�in veritaban�n�z�n bar�nd��� db klas�r� ve y�kleme i�lemlerinin yap�laca�� uploads isimli klas�rlere yazma izni verdi�inizden emin olun,<BR><BR>yazma izinlerinin nas�l verildigi konusunda fikir sahibi degilseniz hosting firman�za bir mail yazarak bu i�lemi tamamlatabilirsiniz.<BR><BR>Kurulumla ilgili akl�n�za tak�lan sorular� <A HREF="http://www.maxinuke.com/forum.asp?action=Control&id=1" target="_blank"><B>Maxinuke.com Forum</B></A> adresinden sorabilir yada mevcut ��z�mleri g�rebilirsiniz.<BR><BR>Haz�rsan�z alt k�s�mdan ilerleyi t�klayarak size rehberlik edecek kur sayfas�n� �al��t�rabilirsiniz.<BR><BR><a href="?eylem=basla"><button onclick="parentElement.click();" CLASS="submit">>> �LERLE VE KURULUMA BA�LA >></button></a><!-- <BR><BR><a href="?eylem=upgrade"><button onclick="parentElement.click();" CLASS="submit">>> �NCEDEN KURULU S�TEM VAR VER� TABANINI AL >></button></a> --></TD>
	</TR>
</TABLE>
<%
	End if
	elseIf Request.QueryString("eylem") = "basla" Then
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.s_name.value == "") {alert("Sitenizin �smi Alan� Bo� !"); return false; }
if (form.s_address.value == "") {alert("Sitenizin Adresi Alan� Bo� !"); return false; }
if (form.s_address.value == "<%=Request.ServerVariables("HTTP_REFERER")%>") {alert("Sitenizin Adres Alan�ndan kur.asp yazisini silin!"); return false; }
if (form.site_slogan.value == "") {alert("Sitenizin Slogani Alan� Bo� !"); return false; }
if (form.s_email.value == "") {alert("Sitenizin E-Mail Adresi Alan� Bo� !"); return false; }
if (form.isim.value == "") {alert("Y�netici Ad ve soyad� Alan� Bo� !"); return false; }
if (form.sehir.value == "0") {alert("Y�netici �ehir Alan� Bo� !"); return false; }
if (form.kul_adi.value == "") {alert("Y�netici Nick Alan� Bo� !"); return false; }
if (form.sifre.value == "") {alert("Y�netici �ifresi Alan� Bo� !"); return false; }
if (form.email.value == "") {alert("Y�netici E-Maili Alan� Bo� !"); return false; }
if (form.g_soru.value == "") {alert("Y�netici Gizli Sorusu Alan� Bo� !"); return false; }
if (form.g_cevap.value == "") {alert("Y�netici Gizli Sorusu Cevabi Alan� Bo� !"); return false; }
return true;
}
</SCRIPT>
<script>
var degerler="1234567890"
	function koduret(plength)
		{
		sabit=''
		for (i=0;i<plength;i++)
		sabit+=degerler.charAt(Math.floor(Math.random()*degerler.length))
		return sabit
		}
	function populateform(enterlength)
		{
		document.kayit.gk.value=koduret(enterlength)
		}
</script>
<form name="kayit" action="?eylem=isle" method="post" onSubmit="return formkontrol(this)">
<input type="hidden" name="uzunluk" value="9">
<input type="hidden" name="gk">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu" style="<%=TableBorder%>">
	<TR>
		<TD align="center" colspan="4"><IMG SRC="Images/kur.png"></TD>
	</TR>
	<tr>
		<td colspan="4" align="center"><hr><b>SITE BILGILERI</b></td>
	</tr>
	<tr>
		<td><b>Sitenizin �smi</b></td>
		<td><input type="text" name="s_name" size="30" class="form1"></td>
		<td><b>Sitenizin Adresi</b></td>
		<td><input type="text" name="s_address" size="30" class="form1" value="<%=Request.ServerVariables("HTTP_REFERER")%>"></td>
	</tr>
	<tr>
		<td><b>Sitenizin Slogan�</b></td>
		<td><input type="text" name="site_slogan" size="30" class="form1" value=""></td>
		<td><b>Sitenizin E-Mail Adresi</b></td>
		<td><input type="text" name="s_email" size="30" class="form1"></td>
	</tr>
	<tr>
		<td><b>Forum Konu Say�s�</b></td>
		<td>
		<select name="f_posts" size="1" class="form1">
		<% For i = 5 TO 50 %>
		<option value="<%=i%>"><%=i%></option>
		<% Next %>
		</select>
		</td>
		<td><b>Forum Cevap Say�s�</b></td>
		<td>
		<select name="f_replies" size="1" class="form1">
		<% For i = 5 TO 50 %>
		<option value="<%=i%>"><%=i%></option>
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td><b>Flood Koruma S�resi</b></td>
		<td>
		<select name="flood_time" size="1" class="form1">
		<% For i = 1 TO 60 %>
		<option value="<%=i%>"><%=i%></option>
		<% Next %>
		</select> Dk.
		</td>
		<td><b>Bir Sayfadaki Makale ve Program Say�s�</b></td>
		<td>
		<select name="s_pacount" size="1" class="form1">
		<% For i = 5 TO 50 %>
		<option value="<%=i%>"><%=i%></option>
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td><b>Mesajla�ma Limiti</b></td>
		<td>
		<select name="s_msglimit" size="1" class="form1">
		<% For i = 10 TO 500 %>
		<option value="<%=i%>"><%=i%></option>
		<% Next %>
		</select>
		</td>
		<td><b>Site LCID (Saat/Tarih Format�)</b></td>
		<td>
		<select name="s_lcid" size="1" class="form1">
		<option value="1055">T�rkiye</option>
		<option value="1033">United States</option>
		<option value="2057">United Kingdom</option>
		<option value="1026">Bulgaristan</option>
		<option value="1030">Danimarka</option>
		<option value="1031">Almanya</option>
		<option value="1053">Isvec</option>
		</select>
		</td>
	</tr>
	<tr>
		<td><b>�yelik Sistemi</b></td>
		<td>
		<select name="onaylama" size="1" class="form1">
		<option value="False" selected>Direk �yelik</option>
		<option value="True">Mail Onayl� �yelik</option>
		</select>
		</td>
		<td><b>Google Adsense Reklam Sistemi</b></td>
		<td>
		<select name="google" size="1" class="form1">
		<option value="True">Evet Olsun</option>
		<option value="False">Hay�r Olmas�n</option>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="4" align="center"><hr><b>SITE SAHIBI BILGILERI</b></td>
	</tr>

	<tr>
		<td><b>Y�netici Ad ve Soyad� : </b></td>
		<td><input type="text" name="isim" size="30" class="form1"></td>
		<td><b>Y�netici �ehri : </b></td>
		<td>
		<SELECT name="sehir" class="form1" >
		<OPTION value=0 selected>- L�tfen �ehir Se�imi Yap�n�z. -</OPTION>
		<OPTION value="Adana">Adana</OPTION>
		<OPTION value="Ad�yaman">Ad�yaman</OPTION>
		<OPTION value="Afyon">Afyon</OPTION>
		<OPTION value="A�r�">A�r�</OPTION>
		<OPTION value="Aksaray">Aksaray</OPTION>
		<OPTION value="Amasya">Amasya</OPTION>
		<OPTION value="Ankara">Ankara</OPTION>
		<OPTION value="Antalya">Antalya</OPTION>
		<OPTION value="Ardahan">Ardahan</OPTION>
		<OPTION value="Artvin">Artvin</OPTION>
		<OPTION value="Ayd�n">Ayd�n</OPTION>
		<OPTION value="Bal�kesir">Bal�kesir</OPTION>
		<OPTION value="Bart�n">Bart�n</OPTION>
		<OPTION value="Batman">Batman</OPTION>
		<OPTION value="Bayburt">Bayburt</OPTION>
		<OPTION value="Bilecik">Bilecik</OPTION>
		<OPTION value="Bing�l">Bing�l</OPTION>
		<OPTION value="Bitlis">Bitlis</OPTION>
		<OPTION value="Bolu">Bolu</OPTION>
		<OPTION value="Burdur">Burdur</OPTION>
		<OPTION value="Bursa">Bursa</OPTION>
		<OPTION value="�anakkale">�anakkale</OPTION>
		<OPTION value="�ank�r�">�ank�r�</OPTION>
		<OPTION value="�orum">�orum</OPTION>
		<OPTION value="Denizli">Denizli</OPTION>
		<OPTION value="Diyarbak�r">Diyarbak�r</OPTION>
		<OPTION value="D�zce">D�zce</OPTION>
		<OPTION value="Edirne">Edirne</OPTION>
		<OPTION value="Elaz��">Elaz��</OPTION>
		<OPTION value="Erzincan">Erzincan</OPTION>
		<OPTION value="Erzurum">Erzurum</OPTION>
		<OPTION value="Eski�ehir">Eski�ehir</OPTION>
		<OPTION value="Gaziantep">Gaziantep</OPTION>
		<OPTION value="Gazimagosa">Gazimagosa</OPTION>
		<OPTION value="Giresun">Giresun</OPTION>
		<OPTION value="Girne">Girne</OPTION>
		<OPTION value="G�m��hane">G�m��hane</OPTION>
		<OPTION value="G�zelyurt">G�zelyurt</OPTION>
		<OPTION value="Hakkari">Hakkari</OPTION>
		<OPTION value="Hatay">Hatay</OPTION>
		<OPTION value="I�d�r">I�d�r</OPTION>
		<OPTION value="Isparta">Isparta</OPTION>
		<OPTION value="Mersin">��el (Mersin)</OPTION>
		<OPTION value="�skele">�skele</OPTION>
		<OPTION value="�stanbul-Anadolu">�stanbul-Anadolu</OPTION>
		<OPTION value="�stanbul-Avrupa">�stanbul-Avrupa</OPTION>
		<OPTION value="�zmir">�zmir</OPTION>
		<OPTION value="Kahramanmara�">Kahramanmara�</OPTION>
		<OPTION value="Karab�k">Karab�k</OPTION>
		<OPTION value="Karaman">Karaman</OPTION>
		<OPTION value="Kars">Kars</OPTION>
		<OPTION value="Kastamonu">Kastamonu</OPTION>
		<OPTION value="Kayseri">Kayseri</OPTION>
		<OPTION value="K�r�kkale">K�r�kkale</OPTION>
		<OPTION value="K�rklareli">K�rklareli</OPTION>
		<OPTION value="K�r�ehir">K�r�ehir</OPTION>
		<OPTION value="Kilis">Kilis</OPTION>
		<OPTION value="Kocaeli">Kocaeli</OPTION>
		<OPTION value="Konya">Konya</OPTION>
		<OPTION value="K�tahya">K�tahya</OPTION>
		<OPTION value="Lefkosa">Lefkosa</OPTION>
		<OPTION value="Malatya">Malatya</OPTION>
		<OPTION value="Manisa">Manisa</OPTION>
		<OPTION value="Mardin">Mardin</OPTION>
		<OPTION value="Mu�la">Mu�la</OPTION>
		<OPTION value="Mu�">Mu�</OPTION>
		<OPTION value="Nev�ehir">Nev�ehir</OPTION>
		<OPTION value="Ni�de">Ni�de</OPTION>
		<OPTION value="Ordu">Ordu</OPTION>
		<OPTION value="Osmaniye">Osmaniye</OPTION>
		<OPTION value="Rize">Rize</OPTION>
		<OPTION value="Sakarya">Sakarya</OPTION>
		<OPTION value="Samsun">Samsun</OPTION>
		<OPTION value="Siirt">Siirt</OPTION>
		<OPTION value="Sinop">Sinop</OPTION>
		<OPTION value="Sivas">Sivas</OPTION>
		<OPTION value="�anl�urfa">�anl�urfa</OPTION>
		<OPTION value="��rnak">��rnak</OPTION>
		<OPTION value="Tekirda�">Tekirda�</OPTION>
		<OPTION value="Tokat">Tokat</OPTION>
		<OPTION value="Trabzon">Trabzon</OPTION>
		<OPTION value="Tunceli">Tunceli</OPTION>
		<OPTION value="U�ak">U�ak</OPTION>
		<OPTION value="Van">Van</OPTION>
		<OPTION value="Yalova">Yalova</OPTION>
		<OPTION value="Yozgat">Yozgat</OPTION>
		<OPTION value="Zonguldak">Zonguldak</OPTION>
		<OPTION value="Yurtdisi">Yurtdisi</OPTION>
		</SELECT>
		</td>
	</tr>
	<tr>
		<td><b>Y�netici Nicki : </b></td>
		<td><input type="text" name="kul_adi" size="30" class="form1"></td>
		<td><b>Y�netici �ifresi : </b></td>
		<td><input type="text" name="sifre" size="30" class="form1"></td>
	</tr>
	<tr>
		<td><b>Y�netici E-Maili : </b></td>
		<td><input type="text" name="email" size="30" class="form1"></td>
		<td><b>Y�netici Gizli Sorusu : </b></td>
		<td><input type="text" name="g_soru" size="30" class="form1"></td>
	</tr>
	<tr>
		<td><b>Y�netici Gizli Sorusu Cevabi : </b></td>
		<td><input type="text" name="g_cevap" size="30" class="form1"></td>
		<td><b>Dogum Tarihi : </b></td><td>
		<select name="yas_1" size="1" class="form1">
		<% For i = 1 TO 31 %>
		<option value="<%=i%>"><%=i%></option>"
		<% Next %>
		</select>
		.
		<select name="yas_2" size="1" class="form1">
		<% For i = 1 TO 12 %>
		<option value="<%=i%>"><%=i%></option>"
		<% Next %>
		</select>
		.
		<select name="yas_3" size="1" class="form1">
		<% For i = 1910 TO 2010 %>
		<option value="<%=i%>"><%=i%></option>"
		<% Next %>
		</select>
		</td>
	  </tr>
	<tr>
		<td colspan="4" align="center"><hr><b>ANA SAYFADA OLACAKLAR</b></td>
	</tr>
	<tr>
		<td><b>Son Dosyalar</b></td>
		<td>
		<select name="sd" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Son Makaleler</b></td>
		<td>
		<select name="sm" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>	
	<tr>
		<td><b>Son �yeler</b></td>
		<td>
		<select name="su" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Son F�kralar</b></td>
		<td>
		<select name="sf" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>		
	<tr>
		<td><b>Son Ekartlar</b></td>
		<td>
		<select name="se" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Son Forum Ba�l�klar�</b></td>
		<td>
		<select name="sfb" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>	
	<tr>
		<td><b>Son Resimler</b></td>
		<td>
		<select name="sr" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Son Oyunlar</b></td>
		<td>
		<select name="so" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>		
	<tr>
		<td><b>Son V�deolar</b></td>
		<td>
		<select name="sv" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td>-</td>
		<td>-</td>
	</tr>			
	<tr>
		<td><b>Rasgele Dosya</b></td>
		<td>
		<select name="rd" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Rasgele Makale</b></td>
		<td>
		<select name="rm" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>	
	<tr>
		<td><b>Rasgele �ye</b></td>
		<td>
		<select name="ru" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Rasgele F�kra</b></td>
		<td>
		<select name="rf" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>		
	<tr>
		<td><b>Rasgele Ekart</b></td>
		<td>
		<select name="re" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Rasgele Video</b></td>
		<td>
		<select name="rv" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>	
	<tr>
		<td><b>Rasgele Resim</b></td>
		<td>
		<select name="rr" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
		<td><b>Rasgele Oyun</b></td>
		<td>
		<select name="ro" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>
	<tr>
		<td><b>Radyo MP3 Player</b></td>
		<td>
		<select name="rplayer" size="1" class="form1">
		<option value="False" selected>Hayir Olmasin</option>
		<option value="True">Evet Olsun</option>
		</select>
		</td>
		<td><b>Son Dakika Haberleri</b></td>
		<td>
		<select name="son_dakika" size="1" class="form1">
		<option value="False">Hayir Olmasin</option>
		<option value="True" selected>Evet Olsun</option>
		</select>
		</td>
	</tr>	
	<tr>
		<td colspan="4"><TEXTAREA NAME="Keywords" ROWS="10" style="width=100%" class="form1" onFocus="javascript:this.value='';">Sitenizin Google vb. arama motorlar�nda �n s�ralarda ��kmas� i�in tan�t�c� kelimeler; her kelimeyi ; i�areti ile ay�r�n�z.</TEXTAREA></td>
	</tr>
	<tr>
		<td colspan="4"><TEXTAREA NAME="sozlesme" ROWS="10" style="width=100%" class="form1" onFocus="javascript:this.value='';">Site s�zlesmesi metni.</TEXTAREA></td>
	</tr>
	<tr>
		<td colspan="4" align="center"><input type="submit" name="submit" onClick="populateform(this.form.uzunluk.value)" value="Tamam �" class="submit" style="width:100%"></td>
	</tr>
</table>
</form>
<%
elseIf Request.QueryString("eylem") = "isle" Then
isim = duz(Request.Form("isim"))
sehir = duz(Request.Form("sehir"))
kul_adi = duz(Request.Form("kul_adi"))
sifre = duz(Request.Form("sifre"))
sifre = MN_PP(sifre)
email = duz(Request.Form("email"))
g_soru = duz(Request.Form("g_soru"))
g_cevap = duz(Request.Form("g_cevap"))
g_cevap = MN_PP(g_cevap)
yas = Request.Form("yas_1") & "." & Request.Form("yas_2") & "." & Request.Form("yas_3")
gk = duz(Request.Form("gk"))
name = duz(Request.Form("s_name"))
address = duz(Request.Form("s_address"))
site_logo = "images/logo.png"
site_slogan = duz(Request.Form("s_slogan"))
flood_time = duz(Request.Form("flood_time"))
email = duz(Request.Form("s_email"))
f_posts = duz(Request.Form("f_posts"))
f_replies = duz(Request.Form("f_replies"))
pa = duz(Request.Form("s_pacount"))
msgLimit = duz(Request.Form("s_msglimit"))
np_msg = "<font class=td_menu><b>Bu Sayfa Sadece �yeler A��kt�r !!!</b><br><br><div align=left>- <a href=uye_islem.asp?action=new>�ye Ol</a><br>- <a href=uye_islem.asp?action=lostpass&step=1>�ifremi Unuttum !!</a></div></font>"
fpass = "Site"
s_lcid = duz(Request.Form("s_lcid"))
google = duz(Request.Form("google"))
onaylama = duz(Request.Form("onaylama"))
son_dakika = duz(Request.Form("son_dakika"))
sd = duz(Request.Form("sd"))
sm = duz(Request.Form("sm"))
su = duz(Request.Form("su"))
sf = duz(Request.Form("sf"))
se = duz(Request.Form("se"))
sfb = duz(Request.Form("sfb"))
sr = duz(Request.Form("sr"))
so = duz(Request.Form("so"))
sv = duz(Request.Form("sv"))
rd = duz(Request.Form("rd"))
rm = duz(Request.Form("rm"))
ru = duz(Request.Form("ru"))
rf = duz(Request.Form("rf"))
re = duz(Request.Form("re"))
rr = duz(Request.Form("rr"))
ro = duz(Request.Form("ro"))
rv = duz(Request.Form("rv"))
rplayer = duz(Request.Form("rplayer"))
Set adminiekle = Server.CreateObject("ADODB.Recordset")
eaSQL = "Select * FROM MEMBERS"
adminiekle.open eaSQL,Connection,3,3
adminiekle.AddNew
adminiekle("isim") = isim
adminiekle("sehir") = sehir
adminiekle("meslek") = "Y�netici"
adminiekle("kul_adi") = kul_adi
adminiekle("sifre") = sifre
adminiekle("email") = email
adminiekle("g_soru") = g_soru
adminiekle("g_cevap") = g_cevap
adminiekle("cinsiyet") = a
adminiekle("yas") = yas
adminiekle("imza") = "-"
adminiekle("login_sayisi") = 0
adminiekle("uyelik_tarihi") = now()
adminiekle("seviye") = 1
adminiekle("msg_sayisi") = 0
adminiekle("last_ip") = Request.ServerVariables("REMOTE_ADDR")
adminiekle("u_theme") = "a"
adminiekle("imza") = "-"
adminiekle("onay") = True
adminiekle("gk") = gk
adminiekle.Update

Set ayarlariekle = Server.CreateObject("ADODB.RecordSet")
nsSQL = "Select * FROM SETTINGS"
ayarlariekle.open nsSQL,Connection,3,3
ayarlariekle.AddNew
ayarlariekle("site_isim") = duz(Request.Form("s_name"))
ayarlariekle("site_adres") = duz(Request.Form("s_address"))
ayarlariekle("site_logo") = "images/logo.png"
ayarlariekle("site_slogan") = duz(Request.Form("site_slogan"))
ayarlariekle("flood_time") = duz(Request.Form("flood_time"))
ayarlariekle("site_mail") = duz(Request.Form("s_email"))
ayarlariekle("prg_sayisi") = duz(Request.Form("s_pacount"))
ayarlariekle("msg_limit") = duz(Request.Form("s_msglimit"))
ayarlariekle("theme") = "a"
ayarlariekle("np_msg") = "<font class=td_menu><b>Bu Sayfa Sadece �yeler A��kt�r !!!</b><br><br><div align=left>- <a href=uye_islem.asp?action=new>�ye Ol</a><br>- <a href=uye_islem.asp?action=lostpass&step=1>�ifremi Unuttum !!</a></div></font>"
ayarlariekle("fpass") = "Site"
ayarlariekle("s_lcid") = duz(Request.Form("s_lcid"))
ayarlariekle("google") = duz(Request.Form("google"))
ayarlariekle("onaylama") = duz(Request.Form("onaylama"))
ayarlariekle("logo_tipi") = false
ayarlariekle("son_dakika") = duz(Request.Form("son_dakika"))
ayarlariekle("f_posts") = duz(Request.Form("f_posts"))
ayarlariekle("f_replies") = duz(Request.Form("f_replies"))
ayarlariekle("Keywords") = duz(Request.Form("Keywords"))
ayarlariekle("sozlesme") = duz(Request.Form("sozlesme"))
ayarlariekle("hv") = false
ayarlariekle("hs") = "5"
ayarlariekle("sd") = duz(Request.Form("sd"))
ayarlariekle("sm") = duz(Request.Form("sm"))
ayarlariekle("su") = duz(Request.Form("su"))
ayarlariekle("sf") = duz(Request.Form("sf"))
ayarlariekle("se") = duz(Request.Form("se"))
ayarlariekle("sfb") = duz(Request.Form("sfb"))
ayarlariekle("sr") = duz(Request.Form("sr"))
ayarlariekle("so") = duz(Request.Form("so"))
ayarlariekle("sv") = duz(Request.Form("sv"))
ayarlariekle("rd") = duz(Request.Form("rd"))
ayarlariekle("rm") = duz(Request.Form("rm"))
ayarlariekle("ru") = duz(Request.Form("ru"))
ayarlariekle("rf") = duz(Request.Form("rf"))
ayarlariekle("re") = duz(Request.Form("re"))
ayarlariekle("rr") = duz(Request.Form("rr"))
ayarlariekle("ro") = duz(Request.Form("ro"))
ayarlariekle("rv") = duz(Request.Form("rv"))
ayarlariekle("rplayer") = duz(Request.Form("rplayer"))
ayarlariekle.Update
adminiekle.Close
Set adminiekle = Nothing
ayarlariekle.Close
Set ayarlariekle = Nothing
Connection.Close
Set Connection = Nothing
Response.redirect "?eylem=ok"

elseIf Request.QueryString("eylem") = "ok" Then
session("kuruldumu")="he"
%>
<TABLE CLASS="td_menu" width="100%">
<TR>
	<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
</TR>
<TR>
	<TD align="center"><B>TEBR�KLER</B><BR><BR>Art�k sizinde bir siteniz var.<br><br>Art�k Sitenize �zellikler Ekleyebilir ve �yeler toplayabilirsiniz.<BR><BR>Sitenizin g�venli�i i�in ana dizininizde bulunan <B>KUR.ASP</B> dosyas�n� silmeyi unutmay�n.<BR><BR>Sitenizle ilgili detayl� ayr�nt�lar� d�zenleyebilece�iniz sayfalar� Y�netim panelinizde bulabilirsiniz.<BR><BR>
	<a href="default.asp"><button onclick="parentElement.click();" CLASS="submit">S�TEM� G�STER</button></a><BR><BR>
<!-- 
	<a href="http://www.maxinuke.com/Programs.asp?action=programs&catid=1"><button onclick="parentElement.click();" CLASS="submit">EKLEYEB�LECE��M BLOKLARI G�REY�M</button></a><BR><BR>
	<a href="http://www.maxinuke.com/Programs.asp?action=programs&catid=2"><button onclick="parentElement.click();" CLASS="submit">EKLEYEB�LECE��M MOD�LLER� G�REY�M</button></a><BR><BR>	
	<a href="http://www.maxinuke.com/moduller.asp?mi=15"><button onclick="parentElement.click();" CLASS="submit">S�TEM� MERKEZE BELE�E KAYIT EDEY�M</button></a><BR><BR>
 -->
	</TD>
</TR>
</TABLE>
<%elseIf Request.QueryString("eylem") = "kurulmus" Then%>
<TABLE CLASS="td_menu" width="100%">
<TR>
	<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
</TR>
<TR>
	<TD align="center"><BR><BR><B><%If session("kuruldumu")="" Then%>Maxinuke <%=v%> kurulum uygulamas�na ho�geldiniz,<BR><BR>k�sa ve basit ad�mlarla sitemizi kuracak ve hemen yay�nlamaya ba�layaca��z,<BR><BR>kuruluma ba�lamak i�in <A HREF="kur.asp">BURAYI</A> t�klay�n<%else%>Merhaba, e�er sistemi daha �nceden kurduysan�z g�venli�iniz i�in server�n�zdan kur.asp dosyas�n� siliniz ve daha sonra <A HREF="default.asp">BURAYI</A> t�klay�n�z.<BR><BR>Silim i�lemi yap�lmad�k�a sitenize eri�emezsiniz. Silimi yaptiktan sonra tekrar deneyebilirsiniz.<%End if%></B></TD>
</TR>
</TABLE>
<%elseIf Request.QueryString("eylem") = "upgrade" Then%>
<TABLE CLASS="td_menu" width="100%">
<TR>
	<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
</TR>
<TR>
	<TD align="center"><BR><BR>Upgrade �zelli�ini kullanabilmek i�in serverinizda �nceden hali haz�rda bir maxi nuke s�r�m�n� kullan�yor olman�z gerekmektedir.<BR>Dosyalar�n �e�itlili�ine g�re de�i�imleri teker teker takip edemeyece�inizden i�leme ba�lamadan �nce yeni s�r�me ait t�m dosyalar� ekilerinin �zerine yazd���n�zdan ve eski ile yeni veritabanlar�n�n ayn� klas�rde olduklar�ndan emin olunuz.<BR>Bu i�lem ayr�m yapmaks�z�n eski s�r�m veritaban�n�zdaki her�eyi oldu�u gibi yeni veritaban�n�za g�nderecektir.<BR><BR><a href="?eylem=upgradebasla"><button onclick="parentElement.click();" CLASS="submit">>> AKTARIMA BA�LA >></button></a><BR><BR><a href="kur.asp"><button onclick="parentElement.click();" CLASS="submit"><< ANA SAYFAYA D�N <<</button></a></TD>
</TR>
</TABLE>
<%elseIf Request.QueryString("eylem") = "upgradebasla" Then%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.eski_mdb.value == "") {alert("Bilgilerin alinacagi eski veritabaninin ismini yazmadiniz !"); return false; }
return true;
}
</SCRIPT>
<FORM onSubmit="return formkontrol(this)" name=peno action="?eylem=ubs" method=post>
<TABLE CLASS="td_menu" width="100%">
<TR>
	<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
</TR>
	<tr>
		<td align="center">Maxinuke V <%=v%> sistemine veriler �ekilmeye ba�lanaccak<BR>L�tfen i�lem tamamland� uyar�s�n� alana kadar bu sayfay� kapatmay�n.<BR>��lem s�resi; aktar�lacak verilerinizin �oklu�una ve internet h�z�n�za g�re de�i�kenlik g�sterebilir.</B></td>
	</tr>
	<tr>
		<td align="center"><B>Bilgilerin Yerle�tirilece�i Database :</B> <%=veritabani_adi%>.mdb</td>
	</tr>
	<tr>
		<td align="center"><B>Bilgilerin Al�naca�� Database :<BR>DB/<input type="text" name="eski_mdb" class="form1" size="40">.mdb</B></td>
	</tr>
</tr>
	<tr>
		<td colspan="2" align="center"><INPUT TYPE="submit" style="width=25%" CLASS="submit" value="VERILERI AKTAR"></td>
	</tr>
</table>
</form>
<%
	elseIf Request.QueryString("eylem") = "ubs" Then
	eskisiadi=Request.Form("eski_mdb")
	Set mdbkontrol = Server.CreateObject("Scripting.FileSystemObject")
		IF Not mdbkontrol.FileExists(""&Server.Mappath("db/"&eskisiadi&".mdb")&"") = True Then
		response.write "<div class=td_menu><BR><BR><BR><CENTER>DB klasoru icerisinde Belirttiginiz adda veritabani bulunamadi<BR><BR>Verlerin alinacagi veritabaniyla verilerin yerlestirilecegi veritabaninin ayni klasorde oldugundan veya eski veritabani adini dogru yazdiginizdan emin olun.<BR><BR><BR><input type=button value=GERI GIT onclick=javascript:history.back(); class=submit></CENTER></div>"
		response.end
		End if
	Set mdbkontrol = Nothing
	Set eskibaglanti = Server.CreateObject("ADODB.Connection")
	eskibaglanti.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/"&eskisiadi&".mdb")
'BANLI KISILER
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
				Response.Write "<div align=center class=td_menu>Banlanmis - Yasaklanmis Kisiler Listesi Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing			
'Reklamlar
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
				Response.Write "<div align=center class=td_menu>Reklamlariniz Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
'Bloklar
'				Set aliver = Server.CreateObject("ADODB.RecordSet")
'				aliver.open "Select * FROM bloklar",eskibaglanti,3,3
'				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
'				yerlestir.open "Select * FROM bloklar",Connection,3,3
'				Do Until aliver.EoF
'				yerlestir.AddNew
'				yerlestir("b_name") = aliver("b_name") 
'				yerlestir("b_content") = aliver("b_content")
'				yerlestir("b_align") = aliver("b_align")
'				yerlestir("b_include") = aliver("b_include")
'				yerlestir("b_active") = aliver("b_active")
'				yerlestir("b_queue") = aliver("b_queue")
'				yerlestir.Update
'				aliver.MoveNext
'				Loop
'				Response.Write "<div align=center class=td_menu>Bloklariniz Aktar�ld� !</div>"
'				yerlestir.Close
'				Set yerlestir = Nothing
'				aliver.Close
'				Set aliver = Nothing
'Yorumlar
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
				Response.Write "<div align=center class=td_menu>Yorumlariniz Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing

				
				
				
				
				
				
				
				
				
				
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM makale_kat",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM makale_kat",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name") 
				yerlestir_kat("cat_info") = aliver_kat("cat_info") 
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Makale Kategorileri Aktar�ld� !</div>"
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM makale",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM makale",Connection,3,3
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
				Response.Write "<div align=center class=td_menu>Makaleler Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_kat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing



'Dosya Kategorileri ve dosyalar
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
				Response.Write "<div align=center class=td_menu>Dosya Kategorileri Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>Dosyalar Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
'Ekartlar
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
				Response.Write "<div align=center class=td_menu>E kart GonderilenlerAktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM ekart_kat",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM ekart_kat",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("isim") = aliver("isim") 
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>E kart Kategoriler Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>E kartlar Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM f_abone",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM f_abone",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("uye_id") = aliver("uye_id") 
				yerlestir("kategoriid") = aliver("kategoriid")
				yerlestir("tarih") = aliver("tarih")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Forum Abonelikleri Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Forum Ana Kategorileri Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>Forum Kategor�ler� Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>Forum Mesajlar� Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing		
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
				Response.Write "<div align=center class=td_menu>Sabit Haberler Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>F�kralar Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>F�kra Kategorileri Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM fo",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM fo",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("baslik") = aliver("baslik") 
				yerlestir("metin") = aliver("metin")
				yerlestir("onay") = aliver("onay")
				yerlestir("fo_yol") = aliver("fo_yol")
				yerlestir("resim") = aliver("resim")
				yerlestir("puan") = aliver("puan")
				yerlestir("katilimci") = aliver("katilimci")
				yerlestir("oynanma") = aliver("oynanma")
				yerlestir("cat") = aliver("cat")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Flash Oyunlar Aktar�ld� !</div>"
				Set aliver_kat = Server.CreateObject("ADODB.RecordSet")
				aliver_kat.open "Select * FROM fo_cat",eskibaglanti,3,3
				Set yerlestir_kat = Server.CreateObject("ADODB.RecordSet")
				yerlestir_kat.open "Select * FROM fo_cat",Connection,3,3
				Do Until aliver_kat.EoF
				yerlestir_kat.AddNew
				yerlestir_kat("cat_name") = aliver_kat("cat_name")
				yerlestir_kat("cat_info") = aliver_kat("cat_info")
				yerlestir_kat("cat_image") = aliver_kat("cat_image")
				yerlestir_kat.Update
				aliver_kat.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Flash Oyunlar Kategorileri Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Arkadaslar Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Ziyaretci defteri Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				yerlestir("u_online") = aliver("u_online")
				yerlestir("u_browser") = aliver("u_browser")
				yerlestir("u_ws") = aliver("u_ws")
				yerlestir("onay") = aliver("onay")
				yerlestir("gk") = aliver("gk")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Uyeler Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Menu Kategorileri Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Menu Linkleri Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Uyeler Arasi Ic Mesajlar Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Moduller Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
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
				Response.Write "<div align=center class=td_menu>Haber Kategorileri Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>Haberler Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_kat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Duyurular Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Sayfalar Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Anket Yanitlari Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>Anket Sorulari Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Resimler Aktar�ld� !</div>"
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
				Response.Write "<div align=center class=td_menu>Resim Kategorileri Aktar�ld� !</div>"
				yerlestir_kat.Close
				Set yerlestir_kat = Nothing
				aliver_kat.Close
				Set aliver_cat = Nothing
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				yerlestir("sozlesme") = aliver("sozlesme")
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
				yerlestir("rplayer") = true
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Ayarlar Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM SWEARWORDS",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM SWEARWORDS",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("s_text") = aliver("s_text") 
				yerlestir("s_value") = aliver("s_value")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Argo korumasi Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM tb",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM tb",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("olay") = aliver("olay") 
				yerlestir("ayi") = aliver("ayi")
				yerlestir("gunu") = aliver("gunu")
				yerlestir("yili") = aliver("yili")
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Tarihte Bugun Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
'				Set aliver = Server.CreateObject("ADODB.RecordSet")
'				aliver.open "Select * FROM THEMES",eskibaglanti,3,3
'				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
'				yerlestir.open "Select * FROM THEMES",Connection,3,3
'				Do Until aliver.EoF
'				yerlestir.AddNew
'				yerlestir("t_name") = aliver("t_name") 
'				yerlestir("t_dir") = aliver("t_dir")
'				yerlestir("t_active") = aliver("t_active")
'				yerlestir.Update
'				aliver.MoveNext
'				Loop
'				Response.Write "<div align=center class=td_menu>Temalar Aktar�ld� !</div>"
'				yerlestir.Close
'				Set yerlestir = Nothing
'				aliver.Close
'				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM video_cat",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM video_cat",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("isim") = aliver("isim") 
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Video Kategorileri Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
				Set aliver = Server.CreateObject("ADODB.RecordSet")
				aliver.open "Select * FROM videolar",eskibaglanti,3,3
				Set yerlestir = Server.CreateObject("ADODB.RecordSet")
				yerlestir.open "Select * FROM videolar",Connection,3,3
				Do Until aliver.EoF
				yerlestir.AddNew
				yerlestir("katid") = aliver("katid") 
				yerlestir("buyuk") = aliver("buyuk") 
				yerlestir("yol") = aliver("yol") 
				yerlestir("tarih") = aliver("tarih") 
				yerlestir("onay") = aliver("onay") 
				yerlestir("hit") = aliver("hit") 
				yerlestir("gonderen") = aliver("gonderen") 
				yerlestir("puan") = aliver("puan") 
				yerlestir("katilimci") = aliver("katilimci") 
				yerlestir.Update
				aliver.MoveNext
				Loop
				Response.Write "<div align=center class=td_menu>Videolar Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
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
				Response.Write "<div align=center class=td_menu>Istatistikler Aktar�ld� !</div>"
				yerlestir.Close
				Set yerlestir = Nothing
				aliver.Close
				Set aliver = Nothing
%>
<meta http-equiv="refresh" content="5;url=?eylem=transferok">
��lem sonu�land�r�l�yor . . .
<%
	eskibaglanti.Close
	Set eskibaglanti = Nothing	
	elseIf Request.QueryString("eylem") = "transferok" then
%>
<TABLE CLASS="td_menu" width="100%">
<TR>
	<TD align="center"><IMG SRC="Images/kur.png"><hr></TD>
</TR>
<TR>
	<TD align="center"><B>TEBR�KLER</B><BR><BR>Tum veriler yeni veritabaniniza basariyla aktarildi.<BR><BR>Sitenizin g�venli�i i�in ana dizininizde bulunan <B>KUR.ASP</B> dosyas�n� silmeyi unutmay�n.<BR><BR>Sitenizle ilgili detayl� ayr�nt�lar� d�zenleyebilece�iniz sayfalar� Y�netim panelinizde bulabilirsiniz.<BR><BR>
	<a href="default.asp"><button onclick="parentElement.click();" CLASS="submit">S�TEM� G�STER</button></a><BR><BR>
	<a href="http://www.maxinuke.com/Programs.asp?action=programs&catid=1"><button onclick="parentElement.click();" CLASS="submit">EKLEYEB�LECE��M BLOKLARI G�REY�M</button></a><BR><BR>
	<a href="http://www.maxinuke.com/Programs.asp?action=programs&catid=2"><button onclick="parentElement.click();" CLASS="submit">EKLEYEB�LECE��M MOD�LLER� G�REY�M</button></a><BR><BR>	
	<a href="http://www.maxinuke.com/moduller.asp?mi=15"><button onclick="parentElement.click();" CLASS="submit">S�TEM� MERKEZE BELE�E KAYIT EDEY�M</button></a></TD>
</TR>
</TABLE>
<%End if%>