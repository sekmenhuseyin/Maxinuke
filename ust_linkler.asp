<table width="100%" height="25" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">
  <tr style="font-weight:bold"> 
	<td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=top_menu_link1%>';"><%=top_menu1%></td>
    <td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
    <td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=top_menu_link2%>';"><%=top_menu2%></td>
    <td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
    <td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='dosya.asp';">Dosyalar</td>
    <td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
    <td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='makale.asp';">Yazý/Makale</td>
    <td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
    <td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=top_menu_link5%>';"><%=top_menu5%></td>
    <td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
    <td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=top_menu_link6%>';"><%=top_menu6%></td>
    <td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
    <td width="10%" align="center" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=top_menu_link7%>';"><%=top_menu7%></td>
	<td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<SCRIPT LANGUAGE=JAVASCRIPT>
function arakont(form) 
{
if (form.q.value == "") {alert("Aranacak kelime belirtmediniz !"); return false; }
return true;
}
</SCRIPT>
<div>
<form action="http://maxinuke.com/search.asp" id="cse-search-box" onSubmit="return arakont(this)" method="get">
    <td width="30%" align="right" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=arama%> : <input type="hidden" name="cx" value="partner-pub-5037075360021369:mmz1bi-snnp" /><input type="hidden" name="cof" value="FORID:11" /><input type="hidden" name="ie" value="windows-1254" /><input type="text" name="q" size="21" class="form1" value="<%=Request("search")%>" /><input type="image" name="sa" src="images/temalar/<%=Session("SiteTheme")%>/ara.gif" border="0" align="absmiddle">&nbsp;
	</td>
</form>
</div>
	<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><A HREF="rss.asp"><IMG SRC="images/rss.png" BORDER="0" ALT="RSS Akislarini Goster"></A></td>  
	<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><a href="javascript:window.external.AddFavorite('<%=sett("site_adres")%>/','<%=sett("site_isim")%>')"><IMG SRC="images/favori.png" BORDER="0" ALT="Favorilere eklemek icin Tiklayin"></a></td>
	<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><a onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('<%=sett("site_adres")%>');"><img border="0" src="images/acilis.gif" alt="Açýlýþ sayfasý yap"></a></td>
  </tr>
</table>