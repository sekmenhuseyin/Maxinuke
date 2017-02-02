<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Doviz Kurlari Modul Uygulamasi V 2.1										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
my="moduller/doviz_kurlari/"
Public Function VeriAl(strGelen)
Set objVeriAl = Server.CreateObject("Microsoft.XMLHTTP" )
objVeriAl.Open "GET" , strGelen, FALSE
objVeriAl.sEnd
VeriAl = objVeriAl.Responsetext
SET objVeriAl = Nothing
End Function
strAdres = "http://www.tcmb.gov.tr/kurlar/today.html"
strVeri = VeriAL(strAdres)

iDolar=InStr(strVeri,"USD" )
strDolarAlis=Mid(strVeri,iDolar+38,10) 'alis
strDolarSatis=Mid(strVeri,iDolar+51,10) 'satis
strDolarEalis=Mid(strVeri,iDolar+60,14) 'Efektif alis
strDolarESatis=Mid(strVeri,iDolar+73,14) 'Efektif satis

iEuro=InStr(strVeri,"EUR" )
strEuroAlis=Mid(strVeri,iEuro+38,11) 'alis
strEuroSatis=Mid(strVeri,iEuro+50,11) 'satis
strEuroEalis=Mid(strVeri,iEuro+60,14) 'Efektif alis
strEuroESatis=Mid(strVeri,iEuro+73,14) 'Efektif satis

iDkk=InStr(strVeri,"DKK" )
strDkkAlis=Mid(strVeri,iDkk+35,12) 'alis
strDkkSatis=Mid(strVeri,iDkk+49,12) 'satis
strDkkEalis=Mid(strVeri,iDkk+60,14) 'Efektif alis
strDkkESatis=Mid(strVeri,iDkk+73,13) 'Efektif satis

iSek=InStr(strVeri,"SEK" )
strSekAlis=Mid(strVeri,iSek+35,13) 'alis
strSekSatis=Mid(strVeri,iSek+42,13) 'satis
strSekEalis=Mid(strVeri,iSek+60,14) 'Efektif alis
strSekESatis=Mid(strVeri,iSek+73,12) 'Efektif satis


iGbp=InStr(strVeri,"GBP" )
strGbpAlis=Mid(strVeri,iGbp+33,13) 'alis
strGbpSatis=Mid(strVeri,iGbp+42,13) 'satis
strGbpEalis=Mid(strVeri,iGbp+60,14) 'Efektif alis
strGbpESatis=Mid(strVeri,iGbp+73,8) 'Efektif satis

iJpy=InStr(strVeri,"JPY" )
strJpyAlis=Mid(strVeri,iJpy+33,13) 'alis
strJpySatis=Mid(strVeri,iJpy+44,13) 'satis
strJpyEalis=Mid(strVeri,iJpy+60,14) 'Efektif alis
strJpyESatis=Mid(strVeri,iJpy+73,13) 'Efektif satis

iCapraz=InStr(strVeri,"ÇAPRAZ" )
strcaudusd=Mid(strVeri,iCapraz+2431,7) 'AUD - USD
strcdkkusd=Mid(strVeri,iCapraz+2480,26) 'DKK - USD
strcchfusd=Mid(strVeri,iCapraz+2560,10) 'CHF - USD
strcsekusd=Mid(strVeri,iCapraz+2620,12) 'SEK - USD
strcjpyusd=Mid(strVeri,iCapraz+2670,20) 'JPY - USD
strccadusd=Mid(strVeri,iCapraz+2730,20) 'CAD - USD
strcnokusd=Mid(strVeri,iCapraz+2800,12) 'NOK - USD
strcsarusd=Mid(strVeri,iCapraz+2860,14) 'SAR - USD
strcusdeur=Mid(strVeri,iCapraz+2915,26) 'USD - EUR
strcusdgbp=Mid(strVeri,iCapraz+2990,6) 'USD - GBP
strcusdkwd=Mid(strVeri,iCapraz+3049,6) 'USD - KWD
%>
<div align="center">
<!--#include file="../../inc/adsense_3.asp" -->
  <table width="467" border="0">
  <tr>
    <td align="right"><img src="<%=my%>images/sol_ust.gif" width="15" height="14"></td>
    <td background="<%=my%>images/duz_ust.gif">&nbsp;</td>
    <td align="left"><img src="<%=my%>images/sag_ust.gif" width="15" height="14"></td>
  </tr>
  <tr>
    <td background="<%=my%>images/duz_sol.gif" valign="top"></td>
    <td>
	
	
	
	<table width="380" border="0"  class="td_menu"
  <tr>
    <td>&nbsp;</td>
    <td background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td><img src="<%=my%>images/dolar.gif" width="50" height="50" alt="Dolar"></td>
    <td background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td><img src="<%=my%>images/euro.gif" width="50" height="50" alt="Euro"></td>
    <td background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td><img src="<%=my%>images/dkk.gif" width="50" height="50" alt="Danimarka Kronu"></td>
    <td background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td><img src="<%=my%>images/sek.gif" width="50" height="50" alt="Ýsveç Kronu"></td>
    <td background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td><img src="<%=my%>images/ing_sterlin.gif" width="50" height="50" alt="Ýngiliz Sterlini"></td>
    <td background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td><img src="<%=my%>images/japon_yeni.gif" width="50" height="50" alt="Japon Yeni"></td>
  </tr>
  <tr>
    <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
    </tr>
  <tr>
    <td height="40">Alýþ</td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdolaralis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=streuroalis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdkkalis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strsekalis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strgbpalis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strjpyalis%></td>
  </tr>
   <tr>
    <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
    </tr>
  <tr>
    <td height="40">Satýþ</td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdolarsatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=streurosatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdkksatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strseksatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strgbpsatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strjpysatis%></td>
  </tr>
   <tr>
    <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
    </tr>
  <tr>
    <td height="40">Efektif Alýþ </td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdolarealis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=streuroealis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdkkealis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strsekealis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strgbpealis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strjpyealis%></td>
  </tr>
   <tr>
    <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
    </tr>
  <tr>
    <td height="40">Efektif Satýþ </td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdolaresatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=streuroesatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strdkkesatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strsekesatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strgbpesatis%></td>
    <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
    <td height="40"><%=strjpyesatis%></td>
  </tr>
</table>
	

	
	</td>
    <td  align="left" background="<%=my%>images/duz_sag.gif"></td>
  </tr>
  <tr>
    <td><img src="<%=my%>images/sol_alt.gif" width="15" height="14"></td>
    <td background="<%=my%>images/duz_alt.gif">&nbsp;</td>
    <td><img src="<%=my%>images/sag_alt.gif" width="15" height="14"></td>
  </tr>
</table>
<!--#include file="../../inc/adsense_3.asp" -->
<font class="block_title">Ç A P R A Z&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;K U R L A R</font>
  
  <table width="467" border="0" >
    <tr>
      <td align="right"><img src="<%=my%>images/sol_ust.gif" width="15" height="14"></td>
      <td background="<%=my%>images/duz_ust.gif">&nbsp;</td>
      <td align="left"><img src="<%=my%>images/sag_ust.gif" width="15" height="14"></td>
    </tr>
    <tr>
      <td background="<%=my%>images/duz_sol.gif" valign="top"></td>
      <td><table width="420" border="0"  class="td_menu">
          <tr>
            <td height="40">AUD - USD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcaudusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">Avusturalya Dolarý</td>

          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">DKK - USD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcdkkusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">Danimarka Kronu</td>

          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">CHF - USD </td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcchfusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">ÝSVÝÇRE FRANGI</td>

          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">SEK - USD </td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcsekusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">ÝSVEÇ KRONU </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>

          <tr>
            <td height="40">JPY - USD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcjpyusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">JAPON YENÝ </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>

          <tr>
            <td height="40">CAD - USD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strccadusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">KANADA DOLARI </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">NOK - USD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcnokusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">NORVEÇ KRONU </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">SAR - USD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ABD DOLARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcsarusd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">SUUDÝ ARABÝSTAN RÝYALÝ </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">USD - EUR</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 EURO</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcusdeur%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">ABD DOLARI </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">USD - GBP</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 ÝNGÝLÝZ STERLÝNÝ</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcusdgbp%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">ABD DOLARI </td>
          </tr>
          <tr>
            <td colspan="13" background="<%=my%>images/duz_alt.gif">&nbsp;</td>
          </tr>
          <tr>
            <td height="40">USD - KWD</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">1 KUVEYT DÝNARI</td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40"><%=strcusdkwd%></td>
            <td height="40" background="<%=my%>images/duz_sol.gif">&nbsp;</td>
            <td height="40">ABD DOLARI </td>
          </tr>
      </table></td>
      <td  align="left" background="<%=my%>images/duz_sag.gif"></td>
    </tr>
    <tr>
      <td><img src="<%=my%>images/sol_alt.gif" width="15" height="14"></td>
      <td background="<%=my%>images/duz_alt.gif">&nbsp;</td>
      <td><img src="<%=my%>images/sag_alt.gif" width="15" height="14"></td>
    </tr>
  </table>
<!--#include file="../../inc/adsense_3.asp" -->
</div>
<Hr>
Bu Sayfadaki veriler ve deðerler Türkiye Cumhuriyeti Merkez Bankasý Resmi web sitesinden alýnmakta ve Merkez Bankasý kayýtlarý esas alýnarak anlýk güncellenmektedir. Kurlarýn Kaynak Kontrolü için <a href="http://www.tcmb.gov.tr/kurlar/today.html" target="_blank">burayý</a> týklayabilirsiniz.
