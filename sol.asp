<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<SCRIPT language=javascript type=text/javascript>
function gostergizle(elem, simgedegis)
{
if (document.getElementById)
	{ecBlock = document.getElementById(elem);
	if (ecBlock != undefined && ecBlock != null)
		{if (simgedegis)
            {elemImage = document.getElementById(elem + "Image");}
            if (!simgedegis || (elemImage != undefined && elemImage != null))
            {
                if (ecBlock.currentStyle.display == "none" || ecBlock.currentStyle.display == null || ecBlock.currentStyle.display == "")
                {ecBlock.style.display = "block";}

				else if (ecBlock.currentStyle.display == "block")
                {ecBlock.style.display = "none";}
               
            }
        }
    }
}
</SCRIPT>
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
	<TR>
		<TD align="center"><a href="sag.asp" target="sag"><img src="images/logo_admin.gif" border="0"></a></TD>
	</TR>
	<TR>
		<TD bgcolor="<%=content_bg%>">
		» <a href="default.asp" target="sag">Siteyi Gör</a><BR>
		» <a href="Your_Account.asp?op=Logout" target="sag">Çýkýþ Yap</a><BR>	
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('1', true); return false;" id=1Image><b>GENEL</b></A></TD>
	</TR>
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=1 style="DISPLAY: none">
		» <a href="testserver.asp" target="sag">Server Raporu</a><BR>
		» <a href="mdbcompact.asp" target="sag">MDB Yedekle</a><br>
		» <a href="update.asp" target="sag">Güncellemeler</a><BR>
<% IF Session("Level") = "1" Then %>
		» <a href="settings.asp?action=settings" target="sag">Ayarlar</a><br>
<!-- 		» <a href="upgrade.asp" target="sag">DB Upgrade</a><br> -->
		» <a href="sql_exec.asp" target="sag">SQL Kodu Çalýþtýr</a><br>
		» <a href="mdbedit.asp" target="sag">MDB Edit</a><br>		
<%End if%>		
		</TD>
	</TR>
<% IF Session("Level") = "1" Then %>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('15', true); return false;" id=15Image><b>ANKET ÝÞLEMLERÝ</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=15 style="DISPLAY: none">
		» <a href="anket_a.asp?action=all" target="sag">Anketler</a><br>
		» <a href="anket_a.asp?action=New" target="sag">Anket Ekle</a>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('13', true); return false;" id=13Image><b>ARGO KORUMASI</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=13 style="DISPLAY: none">
		» <a href="swearwords.asp" target="sag">Argo Listesi</a><br>
		» <a href="swearwords.asp?action=ekle" target="sag">Argo Ekle</a><BR>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('8', true); return false;" id=8Image><b>BANLAMA ÝÞLEMLERÝ</b></A></TD>
	</TR>
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=8 style="DISPLAY: none">
		» <a href="ipban.asp" target="sag">Engellenmiþleri Listele</a><br>
		» <a href="ipban.asp?op=New" target="sag">Ýp Adresi Engelle</a><br>
		» <a href="ipban.asp?op=NewMember" target="sag">Üye Engelle</a><br>
		</td>
	</TR>	
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('20', true); return false;" id=20Image><b>BLOK ÝÞLEMLERÝ</b></A></TD>
	</TR>	
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=20 style="DISPLAY: none">
		» <a href="bloklar.asp?action=all" target="sag">Tüm Bloklar</a><br>
		» <a href="bloklar.asp?action=New" target="sag">Yeni Blok</a>
		</TD>
	</TR>
<%
End If
IF Session("Level") = "1" OR Session("Level") = "2" Then'Download Editoru
%>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('23', true); return false;" id=23Image><b>DOSYALAR</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=23 style="DISPLAY: none">
		» <a href="dosya_a.asp?action=butundosya" target="sag">Dosyalarý listele</a><br>
		» <a href="dosya_a.asp?action=new_enter" target="sag">Dosya Ekle</a><br>
		» <a href="dosya_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="dosya_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</TD>
	</TR>

<%
End if
IF Session("Level") = "1" OR Session("Level") = "4" Then'Haber Editoru
%>	
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('22', true); return false;" id=22Image><b>DUYURULAR</b></A></TD>
	</TR>
	<tr>
		<TD bgcolor="<%=content_bg%>"><DIV id=22 style="DISPLAY: none">
		» <a href="notices.asp" target="sag">Duyurular</a><br>
		» <a href="notices.asp?x=New" target="sag">Yeni Duyuru Ekle</a><br>
		</TD>
	</TR>
<%
End If
IF Session("Level") = "1" OR Session("Level") = "2" Then'Download Editoru
%>	
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('6', true); return false;" id=6Image><b>FLASH OYUNLAR</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=6 style="DISPLAY: none">
		» <a href="fo_a.asp?action=hepsi" target="sag">Oyunlari listele</a><br>
		» <a href="fo_a.asp?action=new_enter" target="sag">Oyun Ekle</a><br>
		» <a href="fo_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="fo_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</td>
	</tr>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('5', true); return false;" id=5Image><b>E-KART</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=5 style="DISPLAY: none">
		» <a href="ekart_a.asp?action=hepsi" target="sag">E Kartlari listele</a><br>
		» <a href="ekart_a.asp?action=yeni" target="sag">E-Kart Ekle</a><br>
		» <a href="ekart_a.asp?action=hepsi_ap" target="sag">Arka Plan listele</a><br>
		» <a href="ekart_a.asp?action=yeni_ap" target="sag">Arka Plan Ekle</a><br>
		» <a href="ekart_a.asp?action=hepsi_midi" target="sag">Music listele</a><br>
		» <a href="ekart_a.asp?action=yeni_midi" target="sag">Music Ekle</a><br>
		» <a href="ekart_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="ekart_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</td>
	</tr>	
<%
End If
IF Session("Level") = "1" Then
%>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('4', true); return false;" id=4Image><b>FIKRALAR</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=4 style="DISPLAY: none">
		» <a href="fikra_a.asp?action=hepsi" target="sag">Fikralari listele</a><br>
		» <a href="fikra_a.asp?action=new_enter" target="sag">Fikra Ekle</a><br>
		» <a href="fikra_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="fikra_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</td>
	</tr>
<%
End if
IF Session("Level") = "1" OR Session("Level") = "5" Then
%>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('25', true); return false;" id=25Image><b>FORUM</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=25 style="DISPLAY: none">
		» <a href="forum_a.asp?action=ayarlar" target="sag">Forum Ayarlarý</a><br>
		» <a href="forum_a.asp?action=katekle" target="sag">Kategori Ekle</a><br>
		» <a href="forum_a.asp?action=yeniforum" target="sag">Forum Ekle</a><br>
		» <a href="forum.asp" target="sag">Forumu Gör</a><br>
		</TD>
	</TR>
<%
End If

IF Session("Level") = "1" OR Session("Level") = "4" Then'Haber Editoru
%>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('21', true); return false;" id=21Image><b>HABERLER</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=21 style="DISPLAY: none">
		» <a href="news_a.asp?action=AllNews" target="sag">Haberleri listele</a><br>
		» <a href="news_a.asp?action=new_enter" target="sag">Haber Ekle</a><br>
		» <a href="news_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="news_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		» <a href="fixnews.asp?action=AllNews" target="sag">Sabit Haberleri Listele</a><br>
		» <a href="fixnews.asp?action=new_enter" target="sag">Sabit Haber Ekle</a><br>
		</td>
	</tr>
<%
End If
IF Session("Level") = "1" OR Session("Level") = "3" Then'Makale Editörü
%>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('24', true); return false;" id=24Image><b>MAKALELER</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=24 style="DISPLAY: none">
		» <a href="makale_a.asp?action=Allmakale" target="sag">Makaleleri listele</a><br>
		» <a href="makale_a.asp?action=new_enter" target="sag">Makale Ekle</a><br>
		» <a href="makale_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="makale_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</TD>
	</TR>
<%
End If
IF Session("Level") = "1" Then
%>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('16', true); return false;" id=16Image><b>MENÜ KATEGORÝLERÝ</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=16 style="DISPLAY: none">
		» <a href="menu_cats.asp" target="sag">Tüm Kategoriler</a><br>
		» <a href="menu_cats.asp?op=New" target="sag">Yeni Kategori Ekle</a>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('17', true); return false;" id=17Image><b>MENÜ LÝNKLERÝ</b></A></TD>
	</TR>	
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=17 style="DISPLAY: none">
		» <a href="menu.asp" target="sag">Linkleri Listele</a><br>
		» <a href="menu.asp?op=New" target="sag">Yeni Menü Linki</a>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('19', true); return false;" id=19Image><b>MODUL ÝÞLEMLERÝ</b></A></TD>
	</TR>	
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=19 style="DISPLAY: none">
		» <a href="moduller_a.asp?action=all" target="sag">Tüm Modüller</a><br>
		» <a href="moduller_a.asp?action=New" target="sag">Yeni Modül</a>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('player', true); return false;" id=playerImage><b>PLAYER ÝÞLEMLERÝ</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=player style="DISPLAY: none">
		» <a href="player_m_a.asp?action=AllNews" target="sag">Muzikleri listele</a><br>
		» <a href="player_m_a.asp?action=new_enter" target="sag">Muzik Ekle</a><br>
		» <a href="player_r_a.asp?action=AllNews" target="sag">Radyolari listele</a><br>
		» <a href="player_r_a.asp?action=new_enter" target="sag">Radyo Ekle</a><br>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('12', true); return false;" id=12Image><b>REKLAM ÝÞLEMLERÝ</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=12 style="DISPLAY: none">
		» <a href="reklam.asp?eylem=liste" target="sag">Banner Listesi</a><br>
		» <a href="reklam.asp?eylem=" target="sag">Banner Ekle</a>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('3', true); return false;" id=3Image><b>RESIM GALERISI</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=3 style="DISPLAY: none">
		» <a href="rg_a.asp?action=Allrg" target="sag">Resimleri listele</a><br>
		» <a href="rg_a.asp?action=new_enter" target="sag">Resim Ekle</a><br>
		» <a href="rg_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="rg_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</div>
		</td>
	</tr>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('18', true); return false;" id=18Image><b>SAYFA ÝÞLEMLERÝ</b></A></TD>
	</TR>	
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=18 style="DISPLAY: none">
		» <a href="sayfa_a.asp?action=hepsi" target="sag">Sayfalari listele</a><br>
		» <a href="sayfa_a.asp?action=yenisayfa" target="sag">Sayfa Ekle</a><br>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('tb', true); return false;" id=tbImage><b>TARÝHTE BUGÜN</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=tb style="DISPLAY: none">
		» <a href="tb_a.asp?action=hepsi" target="sag">Olaylarý listele</a><br>
		» <a href="tb_a.asp?action=new_enter" target="sag">Olay Ekle</a><br>
		</td>
	</tr>	
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('9', true); return false;" id=9Image><b>TEMA ÝÞLEMLERÝ</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=9 style="DISPLAY: none">
		» <a href="themes.asp" target="sag">Tema Listesi</a><br>
		» <a href="themes.asp?op=New" target="sag">Tema Ekle</a>
		</TD>
	</TR>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('14', true); return false;" id=14Image><b>ÜYELÝK ÝÞLEMLERÝ</b></A></TD>
	</TR>
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=14 style="DISPLAY: none">
		» <a href="uye_islem_a.asp?action=members" target="sag">Üye Listesi</a><br>
		» <a href="uye_islem_a.asp?action=MsgToMembers" target="sag">Toplu Mesaj Gönder</a><br>
		» <a href="uye_islem_a.asp?action=AdminMsgs" target="sag">Toplu Mesajlar</a><br>
		» <a href="uye_islem_a.asp?action=DeleteMessages" target="sag">Mesajlarý Sil</a><br>
		» <a href="uye_islem_a.asp?action=Mail2Members" target="sag">Toplu Mail Gönder</a><br>
<!--
		» <a href="uye_islem_a.asp?action=" target="sag">Toplu Mailler</a><br>
		» <a href="uye_islem_a.asp?action=DeleteMessages" target="sag">Mailleri Sil</a><br>
 -->
		» <a href="uye_islem_a.asp?action=MemberMsgs" target="sag">Mesajlaþmalarý Gör</a><br>
<%If sett("onaylama") = True Then%>
<script language="JavaScript">
function onaysizlarisil(listForm2)
{ 
   listForm2.target="sag"; 
   listForm2.action="";
   var answer = confirm ("Sistemdeki onay bekleyen tum uyelerinizi silmek istediginize emin misiniz ?")
   if (answer)
       return true;
   else
       return false;
} 
</script>
		» <a onClick="return onaysizlarisil(this)" href="uye_islem_a.asp?action=onaysizsil" target="sag">Onaysizlari Sil</a><br>
<%End if%>
		</TD>
	</TR>

	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('videolar', true); return false;" id=videolarImage><b>VIDEOLAR</b></A></TD>
	</TR>		
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=videolar style="DISPLAY: none">
		» <a href="video_a.asp?action=hepsi" target="sag">Videolari listele</a><br>
		» <a href="video_a.asp?action=new_enter" target="sag">Video Ekle</a><br>
		» <a href="video_a.asp?action=Cats" target="sag">Kategorileri Listele</a><br>
		» <a href="video_a.asp?action=Cat_New" target="sag">Kategori Ekle</a><br>
		</td>
	</tr>
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20">&nbsp;&nbsp;<A onclick="javascript:gostergizle('7', true); return false;" id=7Image><b>ZÝYARETÇÝ DEFTERÝ</b></A></TD>
	</TR>
	<TR>
		<TD bgcolor="<%=content_bg%>"><DIV id=7 style="DISPLAY: none">
		» <a href="zd_a.asp" target="sag">Görüþleri Listele</a><br>
		</td>
	</TR>
<%End if%>
</TABLE>
</center>