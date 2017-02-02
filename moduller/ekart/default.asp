<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												E-Kart Modul Uygulamasi												#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%ScriptYolu = ""&sett("site_adres")&"/moduller.asp?mi="&Request.QueryString("mi")&""%>
<table border="0" cellspacing="2" cellpadding="2" align="center" width="100%" class="td_menu" style="<%=TableBorder%>">
<TR>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem="><B>Favoriler</B></a></TD>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=ekartgoster"><B>E-Kart Göster</B></a></TD>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=soneklenenler"><B>Son Eklenenler</B></a></TD>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=kategoriler"><B>Kategoriler</B></a></TD>
</TR>
</table>
<BR>
<table border="0" cellspacing="2" cellpadding="2" align="center" width="100%" class="td_menu" style="<%=TableBorder%>">
<%
IF Request.QueryString("eylem") = "" Then
' ------------------------FAVORI KARTLAR LISTELENIYOR ----------------------------------------------- 
kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
adresbar = server.HTMLencode("kimden=" & kimden & "&kimdene=" & kimdene & "&kime=" & kime & "&kimee=" & kimee)
adresbar = Replace(adresbar, " ", "%20")
%>
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- FAVORILER -=-</B></TD>
</TR>
<% 
set favoriler = Server.CreateObJect("ADODB.RecordSet")
SQL_a = "Select TOP 40 * from ekartlar where tip=1 and onay=true ORDER BY hit DESC"
favoriler.open SQL_a,Connection,1,3
if favoriler.eof then
%>
<TR>
	<TD align="center" style="<%=TableBorder%>"><CENTER><BR><BR><BR><B>Kayitli E-Kart Yok</B><BR><BR><BR><BR><BR></CENTER></TD>
</tr>
<%					
End if
While Not favoriler.EOF
%>
<TR>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=favoriler("katid")%>&id=<%=favoriler("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=favoriler("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=favoriler("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
If Not favoriler.EOF Then
favoriler.MoveNext
End If
If Not favoriler.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=favoriler("katid")%>&id=<%=favoriler("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=favoriler("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=favoriler("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not favoriler.EOF Then
favoriler.MoveNext
End If
If Not favoriler.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=favoriler("katid")%>&id=<%=favoriler("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=favoriler("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=favoriler("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not favoriler.EOF Then
favoriler.MoveNext
End If
If Not favoriler.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=favoriler("katid")%>&id=<%=favoriler("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=favoriler("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=favoriler("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not favoriler.EOF Then
favoriler.MoveNext
End If
Wend
favoriler.Close
Set favoriler = Nothing
end if
' -------------------------------- SON EKLENEN KARTLAR LISTELENIYOR ------------------------------------ --
IF Request.QueryString("eylem") = "soneklenenler" Then
%>
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- EN SON EKLENEN EKARTLAR-=-</B></TD>
</TR>
<% 
set soneklenenkart = Server.CreateObJect("ADODB.RecordSet")
SQL_b = "Select TOP 40 * from ekartlar where tip=1 and onay=true ORDER BY tarih DESC"
soneklenenkart.open SQL_b,Connection,1,3
if soneklenenkart.eof then
%>
<TR>
	<TD align="center" style="<%=TableBorder%>"><CENTER><BR><BR><BR><B>Kayitli E-Kart Yok</B><BR><BR><BR><BR><BR></CENTER></TD>
</tr>
<%					
End if
While Not soneklenenkart.EOF
%>
<TR>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=soneklenenkart("katid")%>&id=<%=soneklenenkart("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=soneklenenkart("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=soneklenenkart("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
If Not soneklenenkart.EOF Then
soneklenenkart.MoveNext
End If
If Not soneklenenkart.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=soneklenenkart("katid")%>&id=<%=soneklenenkart("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=soneklenenkart("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=soneklenenkart("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not soneklenenkart.EOF Then
soneklenenkart.MoveNext
End If
If Not soneklenenkart.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=soneklenenkart("katid")%>&id=<%=soneklenenkart("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=soneklenenkart("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=soneklenenkart("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not soneklenenkart.EOF Then
soneklenenkart.MoveNext
End If
If Not soneklenenkart.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=soneklenenkart("katid")%>&id=<%=soneklenenkart("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=soneklenenkart("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=soneklenenkart("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
</TR>
<%
End If
If Not soneklenenkart.EOF Then
soneklenenkart.MoveNext
End If
Wend
soneklenenkart.Close
Set soneklenenkart = Nothing
end if
' -------------------------------- KATEGORILER LISTELENIYOR ------------------------------------ 
IF Request.QueryString("eylem") = "kategoriler" Then 
%>
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- KATEGORÝLER -=-</B></TD>
</TR>
<%
set kategoriler = Server.CreateObJect("ADODB.RecordSet")
SQL = "Select * FROM ekart_kat ORDER BY isim"
kategoriler.open SQL,Connection,1,3
%>
<%While Not kategoriler.EOF%>
	<TR>
		<TD><A href="?mi=<%=Request.QueryString("mi")%>&eylem=katici&katid=<%=kategoriler("katid")%>&<%=adresbar%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/folder.gif" border="0">&nbsp;<%=kategoriler("isim")%></A></TD>
<%
If Not kategoriler.EOF Then
kategoriler.MoveNext
End If
If Not kategoriler.EOF Then
%>
		<TD><A href="?mi=<%=Request.QueryString("mi")%>&eylem=katici&katid=<%=kategoriler("katid")%>&<%=adresbar%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/folder.gif" border="0">&nbsp;<%=kategoriler("isim")%></A></TD>
<%
End If
If Not kategoriler.EOF Then
kategoriler.MoveNext
End If
If Not kategoriler.EOF Then
%>
		<TD><A href="?mi=<%=Request.QueryString("mi")%>&eylem=katici&katid=<%=kategoriler("katid")%>&<%=adresbar%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/folder.gif" border="0">&nbsp;<%=kategoriler("isim")%></A></TD>
<%
End If
If Not kategoriler.EOF Then
kategoriler.MoveNext
End If
Wend
%>
</tr>
<%
kategoriler.Close
Set kategoriler = Nothing
end if
' --------------------------- KATEGORI ICERIGI GORUNTULENIYOR ----------------------- 
IF Request.QueryString("eylem") = "katici" Then
kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
adresbar = server.HTMLencode("kimden=" & kimden & "&kimdene=" & kimdene & "&kime=" & kime & "&kimee=" & kimee)
adresbar = Replace(adresbar, " ", "%20")
katid = Request.QueryString("katid")
If katid = "" Then
katid = Request.Form("katid")
End If
If katid = "" Then
Response.Redirect "."
End If
set kategorisi = Server.CreateObJect("ADODB.RecordSet")
SQL_l = "SELECT * FROM ekart_kat WHERE katid="&katid&""
kategorisi.open SQL_l,Connection,1,3
%>
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- <%=kategorisi("isim")%> Kategorisi E-Kartlarý -=-</B></TD>
</TR>
<%
kategorisi.Close
Set kategorisi = Nothing
set kategoriici = Server.CreateObJect("ADODB.RecordSet")
SQL_m = "SELECT * from ekartlar WHERE tip=1 and onay=true and katid="&katid&""
kategoriici.open SQL_m,Connection,1,3
if kategoriici.eof then
%>
<TR>
	<TD align="center" style="<%=TableBorder%>"><CENTER><BR><BR><BR><B>Bu kategori Altinda Henuz Kayitli E-Kart Yok</B><BR><BR><BR><BR><BR></CENTER></TD>
</tr>
<%					
End if
While Not kategoriici.EOF
%>
<TR>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=kategoriici("katid")%>&id=<%=kategoriici("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=kategoriici("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=kategoriici("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
If Not kategoriici.EOF Then
kategoriici.MoveNext
End If
If Not kategoriici.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=kategoriici("katid")%>&id=<%=kategoriici("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=kategoriici("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=kategoriici("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not kategoriici.EOF Then
kategoriici.MoveNext
End If
If Not kategoriici.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=kategoriici("katid")%>&id=<%=kategoriici("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=kategoriici("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=kategoriici("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
<%
End If
If Not kategoriici.EOF Then
kategoriici.MoveNext
End If
If Not kategoriici.EOF Then
%>
	<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=olustur&katid=<%=kategoriici("katid")%>&id=<%=kategoriici("id")%>&<%=adresbar%>"><img border="0" src="uploads/image/ekart/<%=kategoriici("yol")%>" alt="Bu resmi E-kart yapmak için týklayýnýz" width="95" height="75"></a><br><a ONCLICK="window.open('uploads/image/ekart/<%=kategoriici("buyuk")%>','peno','top=0,left=0,width=600,height=600,toolbar=no,scrollbars=auto');">Resmi Büyüt</a>
	</TD>
</TR>
<%
End If
If Not kategoriici.EOF Then
kategoriici.MoveNext
End If
Wend
kategoriici.Close
Set kategoriici = Nothing
end if
' -------------------------------- YENI BIR KART OLUSTURULUYOR ------------------------------------
IF Request.QueryString("eylem") = "olustur" Then
%>
	<script language="JavaScript">
	function Validate ( obj ) {
	if ( (obj.kimden.value == "") || (obj.kime.value == "") ) {
		alert ('Lütfen her iki Ad kýsmýný da doldurunuz');
		return false;
	}
	if ( (obj.kimdene.value == "") || (obj.kimee.value == "") ) {
		alert ('Lütfen her iki Email Adresi kýsmýný da doldurunuz');
		return false;
	}
	if ( (obj.baslik.value == "") ) {
		alert ('Lütfen Baþlýk kýsmýný doldurunuz');
		return false;
	}
	return true;
	}
	</script>
	<SCRIPT language=JavaScript>
	var supportsKeys = false
	var maxLength
	function yazilimiti(f)
	{
		supportsKeys = true
		calcCharLeft(f)
	}
	function calcCharLeft(f)
	{
		maxLength = 250;
        if (f.mesaj.value.length > maxLength){
	        f.mesaj.value = f.mesaj.value.substring(0,maxLength)
		    charleft = 0
        } else {
			charleft = maxLength - f.mesaj.value.length
		}
        f.harf.value = charleft
}
</SCRIPT>
<%
id = Request.QueryString("id")
katid = Request.QueryString("katid")
kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
adresbar = server.HTMLencode("kimden=" & kimden & "&kimdene=" & kimdene & "&kime=" & kime & "&kimee=" & kimee)
adresbar = Replace(adresbar, " ", "%20")
if id = "" Then
Response.Redirect "."
End If
set olustur_resim = Server.CreateObJect("ADODB.RecordSet")
SQL_d = "Select * from ekartlar WHERE tip=1 and onay=true and id="&id&""
olustur_resim.open SQL_d,Connection,1,3
%>
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- E KART OLUÞTURMA MENÜSÜ -=-</B></TD>
</TR>
<FORM action="?mi=<%=Request.QueryString("mi")%>&eylem=onizleme&katid=<%=katid%>&id=<%=id%>" method="post" onsubmit="return Validate(this);">
<TR>
	<TD colspan="5" align="center">
	<input type="hidden" name="buyuk" value="<%=olustur_resim("buyuk")%>">
	<img src="uploads/image/ekart/<%=olustur_resim("buyuk")%>">
<%
olustur_resim.Close
Set olustur_resim = Nothing
%>
	</TD>
</TR>
<TR>
	<TD colspan="3" align="center" valign="top">
		<table border="0" width="100%" class="td_menu">
		<TR>
			<TD>Adýnýz</TD>
			<TD><INPUT name="kimden" value="<%=kimden%>" size="25" maxlength="30" class="submit"></TD>
		</TR>
		<TR>
			<TD>Mailiniz</TD>
			<TD><INPUT name="kimdene" value="<%=kimdene%>" size="25" maxlength="30"  class="submit"></TD>
		</TR>
		<TR>
			<TD>Alýcýnýn Adý</TD>
			<TD><INPUT name="kime" value="<%=kime%>" size="25" maxlength="30"  class="submit"></TD>
		</TR>
		<TR>
			<TD>Alýcýnýn Maýlý</TD>
			<TD><INPUT name="kimee" value="<%=kimee%>" size="25" maxlength="30"  class="submit"></TD>
		</TR>
		<TR>
			<TD>Mesajýnýzýn Baþlýðý</TD>
			<TD><input name="baslik" size="25" maxlength="50" class="submit"></TD>
		</TR>
		</TABLE>
	</TD>
	<TD colspan="2" align="center" valign="top">Mesajýnýz<br><textarea rows="6" cols="30" name="mesaj" class="submit" onkeyup=yazilimiti(this.form)></textarea><br><INPUT disabled size="3" value="250" name="harf"> karakter kaldý</TD>
</TR>
<TR>
	<TD colspan="5" align="center" valign="top">
		<TABLE>
			<TR>
				<TD>
					<table cellspacing="1" cellpadding=1 border="0" class="td_menu">
                    <tr>
                      <td colspan="8" align="center">Arkaplan rengi</td>
                    </tr>
					<tr>
                      <td bgcolor="#FFFFFF">&nbsp;</td>
                      <td bgcolor="#FFF0B3">&nbsp;</td>
                      <td bgcolor="#FF9900">&nbsp;</td>
                      <td bgcolor="#FFCACA">&nbsp;</td>
                      <td bgcolor="#F2F200">&nbsp;</td>
                      <td bgcolor="#99CC00">&nbsp;</td>
                      <td bgcolor="#DDFFFF">&nbsp;</td>
                      <td bgcolor="#C0C0C0">&nbsp;</td>
                    </tr>
                    <tr>
                      <td bgcolor="#FFFFFF"><input type="radio" checked value="FFFFFF" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="FFF0B3" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="FF9900" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="FFCACA" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="F2F200" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="99CC00" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="DDFFFF" name="backcolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="C0C0C0" name="backcolor"></td>
                    </tr>
					</table>	
				</TD>
				<TD>
					<table cellspacing="1" cellpadding="1" border="0" class="td_menu">
                    <tr>
                      <td colspan="8" align="center">Kenarlýk Rengi</td>
                    </tr>
					<tr>
                      <td bgcolor="#000000">&nbsp;</td>
                      <td bgcolor="#666666">&nbsp;</td>
                      <td bgcolor="#5C96CF">&nbsp;</td>
                      <td bgcolor="#660099">&nbsp;</td>
                      <td bgcolor="#FF0000">&nbsp;</td>
                      <td bgcolor="#006600">&nbsp;</td>
                      <td bgcolor="#FF0066">&nbsp;</td>
                      <td bgcolor="#FFCC00">&nbsp;</td>
                    </tr>
                    <tr>
                      <td bgcolor="#FFFFFF"><input type="radio" value="000000" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="666666" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" checked value="5C96CF" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="660099" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="FF0000" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="006600" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="FF0066" name="bordercolor"></td>
                      <td bgcolor="#FFFFFF"><input type="radio" value="FFCC00" name="bordercolor"></td>
                    </tr>
					</table>
				</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD colspan="5" align="center" valign="top">
		<table border="0" width="100%" class="td_menu">
			<tr>
				<td colspan="4" align="center">Arka Plan Resmi</td>
			</tr>
			<tr>
<%
set olustur_arkaplan = Server.CreateObJect("ADODB.RecordSet")
SQL_e = "Select * FROM ekartlar where tip=2 and onay=true ORDER BY id desc"
olustur_arkaplan.open SQL_e,Connection,1,3
back=0
Do while not olustur_arkaplan.Eof
%>
				<td width="20%" align="center"><br><img src="uploads/image/ekart/<%=olustur_arkaplan("buyuk")%>" width="55" height="55"><br><input type=radio name="arkaplan_id" value="<%=olustur_arkaplan("id")%>">
				</td>
<%
olustur_arkaplan.MoveNext
back=back+1
if back mod 4=0 then
%>				
			</tr>
<%
end if
Loop
%>
		</table>
	</td>
</tr>
<%
olustur_arkaplan.Close
Set olustur_arkaplan = Nothing
%>
<TR>
	<TD colspan="5" align="center" valign="top">Karta Eklenecek Müzik (midi) Ses Dosyasý :&nbsp;&nbsp;&nbsp;&nbsp;
			<select name="midi_id" class="td_menu">
			<option selected value="">-----------------Yok-----------------</option>
<%
set olustur_midi = Server.CreateObJect("ADODB.RecordSet")
SQL_f = "Select * FROM ekartlar where tip=3 and onay=true ORDER BY yol"
olustur_midi.open SQL_f,Connection,1,3
While Not olustur_midi.EOF
%>
			<option value="<%=olustur_midi("id")%>"><%=olustur_midi("yol")%></option>
<%
If Not olustur_midi.EOF Then
olustur_midi.MoveNext
End If
Wend
olustur_midi.Close
Set olustur_midi = Nothing
%>
			</select>
	</td>
</tr>
<TR>
	<TD colspan="5" align="center" valign="top">
	Karttaki Yazý Tipi : 
	<input type="radio" value="arial" name="font"><font face=arial color="#000000" size=2>Arial</font>&nbsp;&nbsp;&nbsp;
	<input type="radio" value="times new roman" name="font"><font face="times new roman" color="#000000" size="2">Times New Roman</font>&nbsp;&nbsp;&nbsp;
	<input type="radio" value="verdana" name="font"><font face="verdana" color="#000000" size="2">Verdana</font>&nbsp;&nbsp;&nbsp;
	<input type="radio" checked  value="tahoma" name="font"><font face="tahoma" color="#000000" size="2">Tahoma</font>&nbsp;&nbsp;&nbsp;
	<input type="radio" value="helvetica" name="font"><font face="helvetica" color="#000000" size="2">Helvetica</font>
	</td>
</tr>
<TR>
	<TD colspan="5" align="center" valign="top">Kart okunduðunda, size haber verilsin mi? <input type="Radio" name="haberver" value="evet">Evet Haberim olsun&nbsp;&nbsp;&nbsp;&nbsp; <input type="Radio" name="haberver" value="hayir" checked>Hayýr Gerek Yok</td>
</tr>
<TR>
	<TD colspan="5" align="center" valign="top">Hazýrladýðýnýz Kartýn bir kopyasýný size gönderelim mi? <input type="Radio" name="kartinkopyasi" value="evet">Evet Gönderin&nbsp;&nbsp;&nbsp;&nbsp;<input type="Radio" name="kartinkopyasi" value="hayir" checked>Hayýr istemem</td>
</tr>
<TR>
	<TD colspan="5" align="center" valign="top"><input name="submit" type="submit" name="B1" value="ÖNÝZLEMEYI GÖSTER" class="submit"></td>
</tr>
</FORM>
<%
end if
' -------------------------------- ONIZLEME ILE KARTA BAKILIYOR ------------------------------------
IF Request.QueryString("eylem") = "onizleme" Then
id = Request.QueryString("id")
arkaplan_id = Request.QueryString("arkaplan_id")
midi_id = Request.QueryString("midi_id")
if id = "" Then
	Response.Redirect "default.asp"
End If
kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
adresbar = server.HTMLencode("kimden=" & kimden & "&kimdene=" & kimdene & "&kime=" & kime & "&kimee=" & kimee)
adresbar = Replace(adresbar, " ", "%20")
buyuk = Request.Form("buyuk")
backcolor = Request.Form("backcolor")
bordercolor = Request.Form("bordercolor")
'id = Request.Form("id")
arkaplan_id = Request.Form("arkaplan_id")
midi_id = Request.Form("midi_id")

if arkaplan_id <>"" then
set olustur_back = Server.CreateObJect("ADODB.RecordSet")
SQL_g = "Select * FROM ekartlar where tip=2 and id="&arkaplan_id&""
olustur_back.open SQL_g,Connection,1,3
back = "uploads/image/ekart/" & olustur_back("buyuk")
olustur_back.Close
Set olustur_back = Nothing
else
id ="yok"
end if
midi_id = Request.Form("midi_id")
if midi_id <>"" then
set olustur_midi = Server.CreateObJect("ADODB.RecordSet")
SQL_h = "Select * FROM ekartlar where tip=3 and id="&midi_id&""
olustur_midi.open SQL_h,Connection,1,3
midi = "uploads/image/ekart/" & olustur_midi("buyuk")
olustur_midi.Close
Set olustur_midi = Nothing
else
id ="yok"
end if
font = Request.Form("font")
haberver = Request.Form("haberver")
kartinkopyasi = Request.Form("kartinkopyasi")
baslik = Request.Form("baslik")
mesaj = Request.Form("mesaj")
mesaj = Replace(mesaj, chr(34), "'")
mesaj = Replace(mesaj, vbCrLf, "<br>")

if instr(mesaj, " ") = 0 then
Text= mesaj
YeniText=""
i=1
counter=0
while i<=Len(Text)
if Mid(Text,i,4)="<br>" then
YeniText=YeniText & Mid(Text,i,1)
counter=0
elseif counter<15 then
YeniText=YeniText & Mid(Text,i,1)
else
YeniText=YeniText & "<br>" & Mid(Text,i,1)
counter=0
end if
counter=counter+1
i=i+1
wend
mesajDUZ = yenitext
ELSE
mesajDUZ = mesaj
end if

if instr(baslik, " ") = 0 then
Text= baslik
YeniText=""
i=1
counter=0
while i<=Len(Text)
if Mid(Text,i,4)="<br>" then
YeniText=YeniText & Mid(Text,i,1)
counter=0
elseif counter<12 then
YeniText=YeniText & Mid(Text,i,1)
else
YeniText=YeniText & "<br>" & Mid(Text,i,1)
counter=0
end if
counter=counter+1
i=i+1
wend
baslikDUZ = yenitext
ELSE
baslikDUZ = baslik
end if
%>
<bgsound src="<%=midi%>" loop="-1">
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- E KART ÖNÝZLEME -=-</B></TD>
</TR>
<TR>
	<TD colspan="5" align="center">
			<table border="0" width="100%" bgcolor="#<%=bordercolor%>" cellspacing="1" cellpadding="2">
							  <tr>
								<td width="100%" bgcolor="#FFFFFF">
								<form action="?mi=<%=Request.QueryString("mi")%>&eylem=gonder" method="post">
				<table border="0" width="100%" cellpadding="0" cellspacing="4" bgcolor="#<%=backcolor%>" background="<%=back%>">
									<tr>
									  <td width="100%" colspan="2"><font face="<%=font%>" size="3" color="#002973"><b><%=kime%></b></font></td>
									</tr>

									<tr>
									  <td valign="top"><img src="uploads/image/ekart/<%=buyuk%>"></td>
									  <td valign="top">
					<table border="0" width="100%" cellspacing="0" cellpadding="0">
										  <tr>
											<td width="70%" valign="top"></td>
											<td width="30%" align="right" valign="top"><img border="0" src="images/pul1.gif"></td>
										  </tr>
										  <tr>
											<td width="100%" colspan="2"><font face="<%=font%>" size="3" color="#3DC435"><b><%=baslikDUZ%></font></b></td>
										  </tr>

										  <tr>
											<td width="100%" colspan="2"><font face="<%=font%>" size="2"><%=mesajDUZ%></font></td>
										  </tr>

										  <tr>
											<td width="100%" align="right" colspan="2"><font face="<%=font%>" size="2" color="#FF0000"><%=kimden%></font><br>
											  <font face="<%=font%>" size="1" color="#FF0000"><%=kimdene%></font><br></td>
										  </tr>
						</table>
									  </td>
									</tr>

									<tr>
									  <td width="100%" colspan="2"><font face="<%=font%>" size="3" color="#002973">&nbsp;</font></td>
									</tr>
					</table>
								</td>
							  </tr>
			</table>
	
	</TD>
</TR>
<TR>
	<TD colspan="5" align="center"><input type="button" value="DÜZENLE" onclick="javascript:history.back();" class="submit">&nbsp;&nbsp;&nbsp;<input name="submit" type="submit" value="E Karti Gonder" class="submit">
	</TD>
</TR>
    <input type="hidden" name="id" value="<%=id%>">
    <input type="hidden" name="buyuk" value="<%=buyuk%>">
    <input type="hidden" name="backcolor" value="<%=backcolor%>">
    <input type="hidden" name="bordercolor" value="<%=bordercolor%>">
    <input type="hidden" name="arkaplan_id" value="<%=arkaplan_id%>">
    <input type="hidden" name="midi_id" value="<%=midi_id%>">
    <input type="hidden" name="font" value="<%=font%>">
    <input type="hidden" name="kimden" value="<%=kimden%>">
    <input type="hidden" name="kimdene" value="<%=kimdene%>">
    <input type="hidden" name="kime" value="<%=kime%>">
    <input type="hidden" name="kimee" value="<%=kimee%>">
    <input type="hidden" name="baslik" value="<%=baslik%>">
    <input type="hidden" name="mesaj" value="<%=mesaj%>">
    <input type="hidden" name="haberver" value="<%=haberver%>">
    <input type="hidden" name="kartinkopyasi" value="<%=kartinkopyasi%>">
    </td>
  </tr>
</form>
<%
end if
'///////////////////////////////// KARTA ID ATANIP MDB YE KAYDEDILIYOR /////////////////////////////////
IF Request.QueryString("eylem") = "gonder" Then
id = Request.Form("id")
if id = "" Then
Response.Redirect "."
End If
' kartid Oluþturma Bölümü
Function Parola_Olustur( nNoChars, sValidChars )
	Const szDefault = "abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
	Dim nCount
	Dim sRet
	Dim nNumber
	Dim nLength
	Randomize
	If sValidChars = "" Then
		sValidChars = szDefault		
	End If
	nLength = Len( sValidChars )
	
	For nCount = 1 To nNoChars
		nNumber = Int((nLength * Rnd) + 1)
		sRet = sRet & Mid( sValidChars, nNumber, 1 )
	Next
	Parola_Olustur = sRet	
End Function

kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
backcolor = Request.Form("backcolor")
bordercolor = Request.Form("bordercolor")
id = Request.Form("id")
arkaplan_id = Request.Form("arkaplan_id")
midi_id = Request.Form("midi_id")
font = Request.Form("font")
haberver = Request.Form("haberver")
kartinkopyasi = Request.Form("kartinkopyasi")
baslik = Request.Form("baslik")
mesaj = Request.Form("mesaj")

set kartimdbyaz = Server.CreateObJect("ADODB.RecordSet")
SQL_i = "UPDATE ekartlar SET hit=hit+1 WHERE id="&id&""
kartimdbyaz.open SQL_i,Connection,1,3

set gidenkartyaz = Server.CreateObJect("ADODB.RecordSet")
SQL_k = "Select * from ekart_gonderilen WHERE gonid=-1"
gidenkartyaz.open SQL_k,Connection,1,3
gidenkartyaz.AddNew
gidenkartyaz("resid") = id
gidenkartyaz("kimden") = kimden
gidenkartyaz("kimdene") = kimdene
gidenkartyaz("kime") = kime
gidenkartyaz("kimee") = kimee
gidenkartyaz("haberver") = haberver 
gidenkartyaz("baslik") = baslik
gidenkartyaz("mesaj") = mesaj
gidenkartyaz("backcolor") = backcolor
gidenkartyaz("bordercolor") = bordercolor
gidenkartyaz("backid") = arkaplan_id
gidenkartyaz("midid") = midi_id
gidenkartyaz("font") = font
kartid = "c" + Parola_Olustur( 10, "" )
gidenkartyaz("kartid")= kartid
gidenkartyaz.Update
gidenkartyaz.Close
Set gidenkartyaz = Nothing
'///////////////////////////////// E KART ALICIYA GONDERILIYOR /////////////////////////////////
konu = kimden & " sana ozel bir E-kart gönderdi"
MesajBody = kime & "<font face=verdana sýze=2> için özel!<br><br>" & vbCrLf & vbCrLf
MesajBody = MesajBody & kimden & " (" & kimdene & ")" & " , size özel bir E-kart hazirladi.<br><br>" & vbCrLf & vbCrLf
MesajBody = MesajBody & "Bu e-kart 30 gün için saklanacaktir, lütfen zamani geçmeden e-kartinizi aliniz.<br><br>" & vbCrLf & vbCrLf
MesajBody = MesajBody & "E-kartinizi <a href="&ScriptYolu&"&eylem=goruntule&kartid="&kartid&">BURAYA</a> tiklayarak okuyabilirsiniz" & vbCrLf & vbCrLf
MesajBody = MesajBody & "<br><br>veya <b><a href="""&ScriptYolu&"&eylem=ekartgoster"">Burayi</a></b> tiklayarak acilan sayfaya e-kart ID "
MesajBody = MesajBody & "numaranizi yazarak (kopyala-yapistir yapabilirsiniz) e-kartiniza ulasabilirsiniz.<br><br> E - Kart id Numaraniz :<b>" & vbCrLf & vbCrLf & kartid
MesajBody = MesajBody & vbCrLf & vbCrLf & vbCrLf & "</b><br><br>Siz de sevdiklerinize ücretsiz e-kart göndermek istiyorsaniz <A HREF="&ScriptYolu&">"&sett("site_isim")&"</A> sitesini ziyaret edebilirsiniz.</font>" & vbCrLf & vbCrLf
	Set Email = Server.Createobject("CDONTS.NewMail")
	Email.BodyFormat=0
	Email.MailFormat=0
	Email.From = sett("site_mail")
	Email.To = kimee
	Email.Subject = konu
	Email.Body = (MesajBody)
	Email.Send
	Set Email = Nothing
'///////////////////////////////// E KART KOPYASI GONDERENE GIDIYOR /////////////////////////////////
If kartinkopyasi = "evet" then
kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
konu = "Göndermis oldugunuz E-kart ile ilgili"
MesajBody = "<b>E-kart gönderdiginiz kisi :</b> " & kime & " (" & kimee & ")" & vbCrLf & vbCrLf
MesajBody = MesajBody & vbCrLf & vbCrLf & vbCrLf & "<br><br>Gönderdiginiz e-karti görmek istiyorsaniz <a href="&ScriptYolu&"&eylem=goruntule&kartid="&kartid&"&guncelleme=hayir>Burayi</a> tiklayabilirsiniz.<br><br>"& vbCrLf & vbCrLf
MesajBody = MesajBody & vbCrLf & vbCrLf & vbCrLf & "" & sett("site_isim") & " E Kart Servisini seçtiginiz için tesekkür ederiz.<br><br>"& vbCrLf & vbCrLf
kime = kimden
kimee = kimdene
	Set Email = Server.Createobject("CDONTS.NewMail")
	Email.BodyFormat=0
	Email.MailFormat=0
	Email.From = sett("site_mail")
	Email.To = kimee
	Email.Subject = konu
	Email.Body = (MesajBody)
	Email.Send
	Set Email = Nothing
End If
	tesekkur = "?mi="&Request.QueryString("mi")&"&eylem=ok&kimden=" & kimden & "&kimdene=" & kimdene
	tesekkur = Replace(tesekkur, " ", "%20")
	Response.Redirect (tesekkur)
kartimdbyaz.Close
Set kartimdbyaz = Nothing
end if
'///////////////////////////////// E KART GONDERILDI TESEKKUR SAYFASI /////////////////////////////////
IF Request.QueryString("eylem") = "ok" Then
kimden = Request("kimden")
kimdene = Request("kimdene")
kimdene = Replace(kimdene, " ", "")
kime = Request("kime")
kimee = Request("kimee")
kimee = Replace(kimee, " ", "")
adresbar = server.HTMLencode("kimden=" & kimden & "&kimdene=" & kimdene)
adresbar = Replace(adresbar, " ", "%20")
%>
<TR>
	<TD colspan="5" align="center" height="360">Tebrikler !<br><br><br>E-kartýnýz baþarýyla gönderildi.<br><br><br>Yeni E-kartlar göndermek istiyorsanýz <a href="?mi=<%=Request.QueryString("mi")%>&eylem=&<%=adresbar%>">BURAYI</a> týklayýnýz.<br><br></TD>
</TR>
<%
end if
'///////////////////////////////// E - KARTIMI GÖSTER MENUSU /////////////////////////////////
IF Request.QueryString("eylem") = "ekartgoster" Then
%>
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- E - KARTIMI GÖSTER -=-</B></TD>
</TR>
<TR>
	<TD colspan="5" align="center" height="300">Size gönderilen bir e-kart varsa, kartýnýzýn ID numarasýný aþaðýdaki kutuya yazarak görebilirsiniz.<br><br><br><form method="POST" action="?mi=<%=Request.QueryString("mi")%>&eylem=goruntule"><b>Kart ID:&nbsp;&nbsp;</b><input type="text" name="kartid" size="25" class="submit">&nbsp;&nbsp;<input type="submit" name="B1" value="Göster" class="submit" align="middle" border="0"></form></TD>
</TR>
<%
end if
'///////////////////////////////// E KART GORUNTULENIYOR /////////////////////////////////
IF Request.QueryString("eylem") = "goruntule" Then
kartid = Request.QueryString("kartid")
	if kartid = "" Then
	kartid = Request.Form("kartid")
		If kartid = "" OR Len(kartid) < 11 Then
		Response.Redirect "?mi="&Request.QueryString("mi")&"&eylem=hata"
		End If
	End If
set goruntule1 = Server.CreateObJect("ADODB.RecordSet")
SQL_aa = "select * from ekart_gonderilen where kartid='"&kartid&"'"
goruntule1.open SQL_aa,Connection,1,3
	If goruntule1.EOF Then
	Response.Redirect "?mi="&Request.QueryString("mi")&"&eylem=hata"
	End If
'///////////////////////////////// okundu tarihi bilgileri isleniyor /////////////////////////////////
	if Request.QueryString("guncelleme") = "hayir" then
	else
	Set okundu=Server.CreateObject("Adodb.Recordset")
	okundu.ActiveConnection = Connection
	SQL="SELECT * from ekart_gonderilen where kartid='"&kartid&"'"
	okundu.open SQL,Connection,3,2
	okundu.update "okunma",now()
	okundu.close
	set okundu=nothing
	If haberver = "evet" then
	konu = kime & " E-kartinizi aldi"
	MesajBody = kime & " (" & kimee & ")" & " , göndermis oldugunuz E-kartinizi aldi.<br><br><b>" & vbCrLf & vbCrLf
	MesajBody = MesajBody & "Kartin Alinma Tarihi :</b> " & Now & vbCrLf & vbCrLf
	MesajBody = MesajBody & "<br><br>Göndermis oldugunuz E-karti <a href=""http://" & vbCrLf & vbCrLf
	MesajBody = MesajBody & ""&ScriptYolu&"&eylem=goruntule&kartid="&kartid&"&guncelleme=hayir"">" & vbCrLf & vbCrLf
	MesajBody = MesajBody & "BURAYA</a> tiklayarak tekrar gorebilirsiniz.<br><br>" & vbCrLf & vbCrLf
	MesajBody = MesajBody & vbCrLf & vbCrLf & vbCrLf
	MesajBody = MesajBody & vbCrLf & vbCrLf & vbCrLf & sett("site_isim") & " E-Kart Servisini seçtiginiz için tesekkür ederiz." & vbCrLf
	kime = kimden
	kimee = kimdene
	kimden = sett("site_isim")
	kimdene = sett("site_mail")
	Set Email = Server.Createobject("CDONTS.NewMail")
	Email.BodyFormat=0
	Email.MailFormat=0
	Email.From = sett("site_mail")
	'Email.From = kimdene
	Email.To = kimee
	Email.Subject = konu
	Email.Body = (MesajBody)
	Email.Send
	Set Email = Nothing		
	goruntule1("haberver") = Now()
	goruntule1.Update
	End If
	end if
'///////////////////////////////// e karta bakiliyor /////////////////////////////////
resimid = goruntule1("resid")
kimden = goruntule1("kimden")
kimdene = goruntule1("kimdene")
kime = goruntule1("kime")
kimee = goruntule1("kimee")
baslik = goruntule1("baslik")
mesaj = goruntule1("mesaj")
haberver = goruntule1("haberver")
backcolor = goruntule1("backcolor")
bordercolor = goruntule1("bordercolor")
arkaplan_id = goruntule1("backid")
midi_id = goruntule1("midid")
font = goruntule1("font")
'///////////////////////////////// E KART OKUNDU MAILI GIDIYOR /////////////////////////////////
kimden = goruntule1("kimden")
kimdene = goruntule1("kimdene")
kime = goruntule1("kime")
kimee = goruntule1("kimee")
goruntule1.Close
Set goruntule1 = Nothing
Set hangiresim = Connection.Execute("SELECT * FROM ekartlar WHERE tip=1 and onay=true and id="&resimid&"")
buyuk = hangiresim("buyuk")
hangiresim.Close
if id <>"yok" then
Set goruntule1 = Connection.Execute("SELECT * FROM ekartlar where tip=2 and id=" & arkaplan_id)
back = "uploads/image/ekart/" & goruntule1("buyuk")
goruntule1.Close
else
back =""
end if

if id <>"yok" then
Set goruntule1 = Connection.Execute("SELECT * FROM ekartlar where tip=3 and id=" & midi_id)
midi = "uploads/image/ekart/" & goruntule1("buyuk")
goruntule1.Close
else
midi =""
end if

if instr(mesaj, " ") = 0 then
Text= mesaj
YeniText=""
i=1
counter=0
while i<=Len(Text)
if Mid(Text,i,4)="<br>" then
YeniText=YeniText & Mid(Text,i,1)
counter=0
elseif counter<15 then
YeniText=YeniText & Mid(Text,i,1)
else
YeniText=YeniText & "<br>" & Mid(Text,i,1)
counter=0
end if
counter=counter+1
i=i+1
wend
mesajDUZ = yenitext
ELSE
mesajDUZ = mesaj
end if

if instr(baslik, " ") = 0 then
Text= baslik
YeniText=""
i=1
counter=0
while i<=Len(Text)
if Mid(Text,i,4)="<br>" then
YeniText=YeniText & Mid(Text,i,1)
counter=0
elseif counter<12 then
YeniText=YeniText & Mid(Text,i,1)
else
YeniText=YeniText & "<br>" & Mid(Text,i,1)
counter=0
end if
counter=counter+1
i=i+1
wend
baslikDUZ = yenitext
ELSE
baslikDUZ = baslik
end if
%>
<bgsound src="<%=midi%>" loop="-1">
<TR>
	<TD colspan="5" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20" align="center"><B>-=- <%=sett("site_isim")%> E-KART SERVÝSÝ -=-</B></TD>
</TR>
<TR>
	<TD colspan="5" align="center">
		<table border="0" width="100%" cellpadding="0" cellspacing="4" bgcolor="#<%=backcolor%>" background="<%=back%>">
			<tr>
				<td width="100%" colspan="2"><font face="<%=font%>" size="3" color="#002973"><b><%=kime%></b></font></td>
			</tr>
			<tr>
				<td valign="top"><img src="uploads/image/ekart/<%=buyuk%>"></td>
				<td valign="top">
					<table border="0" width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td width="30%" align="right" valign="top" colspan="2"><img border="0" src="images/pul1.gif"></td>
						</tr>
						<tr>
							<td width="100%" colspan="2"><font face="<%=font%>" size="3" color="#3DC435"><b><%=baslikDUZ%></font></b>
							</td>
						</tr>
						<tr>
							<td width="100%" colspan="2"><font face="<%=font%>" size="2"><%=mesajDUZ%></font></td>
						</tr>
						<tr>
							<td width="100%" align="right" colspan="2"><font face="<%=font%>" size="2" color="#FF0000"><%=kimden%></font><br><font face="<%=font%>" size="1" color="#FF0000"><%=kimdene%></font></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
<%
cevapla = server.HTMLencode("kimden=" & kime & "&kimdene=" & kimee & "&kime=" & kimden & "&kimee=" & kimdene)
cevapla = Replace(cevapla, " ", "%20")
%>
	</td>
</tr>
<%
	if Request.QueryString("guncelleme") = "hayir" then
	else
%>
<TR>
	<TD colspan="5" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=&<%=cevapla%>"><button class="submit" onclick="parentElement.click();">CEVAPLA</button></a></TD>
</TR>
<%
	end if
end if
' --------------------------------------------- HATA OLUSTU ------------------------------------------------- 
IF Request.QueryString("eylem") = "hata" Then
%>
<TR>
	<TD colspan="5" align="center" height="280">Üzgünüz - Geçerli bir E-kart ID numarasý girmediniz.<br><br><br><a href="javascript:history.go(-1)">Geri dönmek için týklayýnýz.</a></TD>
</TR>
<%end if%>

</TABLE>