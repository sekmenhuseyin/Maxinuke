<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										ADSL Kota Hesaplama Modul Uygulamasi V 1.2									#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%IF Request.QueryString("eylem") = "" Then%>
<form method="post" action="?mi=<%=Request.QueryString("mi")%>&eylem=hesapla">
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu">
	<TR>
		<TD>Kota a��m� yap�p yapmad���n�z� ��renmek yada bu ay faturam�zda ne kadar �deyece�imizi bilmek i�in �ncelikle <A HREF="http://adslkota.ttnet.net.tr" target="_blank">http://adslkota.ttnet.net.tr</A> adresine giriyor ve kota bilgilerimizi al�yoruz, daha sonra alt k�s�mdaki byte'm�z k�sm�na kulland���m�z byte'� yaz�yoruz ve faturam�z�n durumunu ��reniyorsunuz.<BR><BR></TD>
	</TR>
	<TR>
		<TD align="center">Tarife �e�idiniz<br><BR>
		<select name="tarife" class="form1">
        <option value="1" selected="selected">3 GB</option>
        <option value="2">6 GB</option>
        <option value="3">9 GB</option>
		</select><BR><BR>
		</TD>
	</TR>
	<TR>
		<TD align="center">Byte'm�z<BR><BR><input name="kullanim" onkeyup="sayi(this);" type="text" class="form1"><BR><BR></TD>
	</TR>
	<TR>
		<TD align="center"><input name="hesapla" type="submit" class="submit" id="hesapla" value="Faturam� Hesapla" /></TD>
	</TR>
</TABLE>
</form>
<%
elseIF Request.QueryString("eylem") = "hesapla" Then
tarife=Request.Form("tarife")
kullanilan=Request.Form("kullanim")
kullanilan=Replace(kullanilan,",","")
kullanilan=Replace(kullanilan,".","")
tarife2=3072000000
	if kullanilan = "" then kullanilan=0
	tarifeson=tarife*tarife2
	sinir="0,008872"
		if isnumeric(sinir) then
		sinir=sinir
		else
		sinir=cint(sinir)
		end if
	kullanilansonuc=kullanilan-tarifeson
		if kullanilansonuc<0 then
		sonuctxt="Kota Ge�ilmemis"
		durum=false
		else
		sonuctxt="<font color=red><b>Kotal Ge�ilmis</b></font>"
		durum=true
		end if
	hesapson=(kullanilansonuc/1024000)*sinir
	yer = Instr(hesapson,",")
		if yer>0 then
		hesapson=MID(hesapson,1,YER+2)
		else
		hesapson=hesapson
		end if
	%>
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu">
<TR>
	<TD align=right>Kullanman�z Gereken</TD>
	<TD>: <%=tarifeson%> byte</TD>
</TR>
<TR>
	<TD align=right>Kulland���n�z</TD>
	<TD>: <%=kullanilan%> byte</TD>
</TR>
<TR>
	<TD align=right>Sonu�</TD>
	<TD>: <%=sonuctxt%></TD>
</TR>
	<%if durum=true then%>
<TR>
	<TD align=right>�demeniz gereken tutar</TD>
	<TD>: <%=hesapson%> YTL.</TD>
</TR>
	<%end if%>
</TABLE>
<%End if%>