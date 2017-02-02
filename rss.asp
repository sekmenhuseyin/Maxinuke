<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
%>
<table border="0" cellpadding="3" cellspacing="3" width="100%" style="<%=TableBorder%>" class="td_menu" bgcolor="<%=content_bg%>">
	<tr>
		<td  class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=sett("site_isim")%> RSS SERVÝSÝ</B></CENTER></td>
	</tr>
	<tr>
		<td valign="top"><IMG SRC="images/rss2.png" align=left>Artýk <%=sett("site_isim")%> RSS Servisi ile de <%=sett("site_isim")%>'in sýk sýk deðiþen içeriðini, son dakika geliþmelerini bilgisayarýnýzdan anýnda takip edebileceksiniz. Bunun için yapmanýz gereken tek þey, aþaðýdaki bölüm listesinden ilginizi çekenleri RSS okuyucu programýnýza kayýt etmek. Aþaðýda kullandýðýnýz iþletim sisteminizde çalýþacak RSS okuyucu programlardan birini seçebilirsiniz. RSS, içerik daðýtýmýný basitleþtiren XML (eXtensible Markup Language) tabanlý bir formattýr. Eðer Windows iþletim sistemi kullanýyorsanýz <%=sett("site_isim")%> RSS servisinden yaralanmak için <A href="http://ftp.e-kolay.net/haberci/E-kolayRSS.exe"><B>EKOLAY RSS okuyucuyu</B></A> kullanabilirsiniz. Farkli Ýþletim sistemleri icin RSS programlarýni burdan indirebilirsiniz. (<A href="http://directory.google.com/Top/Reference/Libraries/Library_and_Information_Science/Technical_Services/Cataloguing/Metadata/RDF/Applications/RSS/News_Readers/Windows/" target=_blank>Windows</A> - <A href="http://directory.google.com/Top/Reference/Libraries/Library_and_Information_Science/Technical_Services/Cataloguing/Metadata/RDF/Applications/RSS/News_Readers/Mac_OS" target=_blank>Mac OS</A> - <A href="http://directory.google.com/Top/Reference/Libraries/Library_and_Information_Science/Technical_Services/Cataloguing/Metadata/RDF/Applications/RSS/News_Readers/Handhelds" target=_blank>Cep Bilgisayarlarý</A> - <A href="http://directory.google.com/Top/Reference/Libraries/Library_and_Information_Science/Technical_Services/Cataloguing/Metadata/RDF/Applications/RSS/News_Readers/Java" target=_blank>Java</A> - <A href="http://directory.google.com/Top/Reference/Libraries/Library_and_Information_Science/Technical_Services/Cataloguing/Metadata/RDF/Applications/RSS/News_Readers/Linux" target=_blank>Linux</A> - <A href="http://directory.google.com/Top/Reference/Libraries/Library_and_Information_Science/Technical_Services/Cataloguing/Metadata/RDF/Applications/RSS/News_Readers/Web_Based" target=_blank>Web Tabanlý</A>)</td>
	</tr>
	<tr>
		<td align="center">
			<table border="0" cellpadding="2" cellspacing="0" class="td_menu" bgcolor="<%=content_bg%>">
				<tr>
					<td align="right"><A HREF="rss_haber.asp"><B>Haberler</B></A></td>
					<td>&nbsp;&nbsp;<A HREF="rss_haber.asp"><A HREF="rss_haber.asp"><%=sett("site_adres")%>/rss_haber.asp</A></td>
				</tr>
				<tr>
					<td align="right"><A HREF="rss_dosya.asp"><B>Dosyalar</B></A></td>
					<td>&nbsp;&nbsp;<A HREF="rss_dosya.asp"><A HREF="rss_dosya.asp"><%=sett("site_adres")%>/rss_dosya.asp</A></td>
				</tr>
				<tr>
					<td align="right"><A HREF="rss_makale.asp"><B>Makaleler</B></A></td>
					<td>&nbsp;&nbsp;<A HREF="rss_makale.asp"><A HREF="rss_makale.asp"><%=sett("site_adres")%>/rss_makale.asp</A></td>
				</TR>
				<tr>
					<td align="right"><A HREF="rss_forum.asp"><B>Forum</B></A></td>
					<td>&nbsp;&nbsp;<A HREF="rss_forum.asp"><A HREF="rss_forum.asp"><%=sett("site_adres")%>/rss_forum.asp</A></td>
				</TR>
			</TABLE>		
		</td>
	</tr>
</table>
<%
call ORTA2
call ALT
%>