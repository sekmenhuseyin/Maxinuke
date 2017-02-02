<%
set son20haber = server.createobject("adodb.recordset")
SQL = "Select * from NEWS where onay=true order by hid desc"
son20haber.open SQL,Connection,1,3
if son20haber.eof then
response.Write "<table width=100% height=70 border=0 cellspacing=0 cellpadding=0 bgcolor="&content_bg&" style="&TableBorder&" class=td_menu><tr><td height=25 colspan=2 align=center background=images/temalar/"&Session(SiteTheme)&"/menu_bg.gif>"&sett("site_isim")&" web sayfamiz yayýn hayatýna baþladý,<BR><BR> þu anda boþ olan tüm alanlar yakýn zamanda güncellenecetir.</td></tr></table>"
else
ilkhaber=1
Do While ilkhaber < 21 and Not son20haber.eof
if ilkhaber=1 then
	hid_1=son20haber("hid")
	hbrbaslik1=son20haber("baslik")
	hbrresim1=son20haber("resim")
	hbrmetin1=Left(son20haber("metin"),160)
elseif ilkhaber=2 then
	hid_2=son20haber("hid")
	hbrbaslik2=son20haber("baslik")
	hbrresim2=son20haber("resim")
	hbrmetin2=Left(son20haber("metin"),160)
elseif ilkhaber=3 then
	hid_3=son20haber("hid")
	hbrbaslik3=son20haber("baslik")
	hbrresim3=son20haber("resim")
	hbrmetin3=Left(son20haber("metin"),160)
elseif ilkhaber=4 then
	hid_4=son20haber("hid")
	hbrbaslik4=son20haber("baslik")
	hbrresim4=son20haber("resim")
	hbrmetin4=Left(son20haber("metin"),160)
elseif ilkhaber=5 then
	hid_5=son20haber("hid")
	hbrbaslik5=son20haber("baslik")
	hbrresim5=son20haber("resim")
	hbrmetin5=Left(son20haber("metin"),160)
elseif ilkhaber=6 then
	hid_6=son20haber("hid")
	hbrbaslik6=son20haber("baslik")
	hbrresim6=son20haber("resim")
	hbrmetin6=Left(son20haber("metin"),160)
elseif ilkhaber=7 then
	hid_7=son20haber("hid")
	hbrbaslik7=son20haber("baslik")
	hbrresim7=son20haber("resim")
	hbrmetin7=Left(son20haber("metin"),160)
elseif ilkhaber=8 then
	hid_8=son20haber("hid")
	hbrbaslik8=son20haber("baslik")
	hbrresim8=son20haber("resim")
	hbrmetin8=Left(son20haber("metin"),160)
elseif ilkhaber=9 then
	hid_9=son20haber("hid")
	hbrbaslik9=son20haber("baslik")
	hbrresim9=son20haber("resim")
	hbrmetin9=Left(son20haber("metin"),160)
elseif ilkhaber=10 then
	hid_10=son20haber("hid")
	hbrbaslik10=son20haber("baslik")
	hbrresim10=son20haber("resim")
	hbrmetin10=Left(son20haber("metin"),160)
elseif ilkhaber=11 then
	hid_11=son20haber("hid")
	hbrbaslik11=son20haber("baslik")
	hbrresim11=son20haber("resim")
	hbrmetin11=Left(son20haber("metin"),160)
elseif ilkhaber=12 then
	hid_12=son20haber("hid")
	hbrbaslik12=son20haber("baslik")
	hbrresim12=son20haber("resim")
	hbrmetin12=Left(son20haber("metin"),160)
elseif ilkhaber=13 then
	hid_13=son20haber("hid")
	hbrbaslik13=son20haber("baslik")
	hbrresim13=son20haber("resim")
	hbrmetin13=Left(son20haber("metin"),160)
elseif ilkhaber=14 then
	hid_14=son20haber("hid")
	hbrbaslik14=son20haber("baslik")
	hbrresim14=son20haber("resim")
	hbrmetin14=Left(son20haber("metin"),160)
elseif ilkhaber=15 then
	hid_15=son20haber("hid")
	hbrbaslik15=son20haber("baslik")
	hbrresim15=son20haber("resim")
	hbrmetin15=Left(son20haber("metin"),160)
elseif ilkhaber=16 then
	hid_16=son20haber("hid")
	hbrbaslik16=son20haber("baslik")
	hbrresim16=son20haber("resim")
	hbrmetin16=Left(son20haber("metin"),160)
elseif ilkhaber=17 then
	hid_17=son20haber("hid")
	hbrbaslik17=son20haber("baslik")
	hbrresim17=son20haber("resim")
	hbrmetin17=Left(son20haber("metin"),160)
elseif ilkhaber=18 then
	hid_18=son20haber("hid")
	hbrbaslik18=son20haber("baslik")
	hbrresim18=son20haber("resim")
	hbrmetin18=Left(son20haber("metin"),160)
elseif ilkhaber=19 then
	hid_19=son20haber("hid")
	hbrbaslik19=son20haber("baslik")
	hbrresim19=son20haber("resim")
	hbrmetin19=Left(son20haber("metin"),160)
elseif ilkhaber=20 then
	hid_20=son20haber("hid")
	hbrbaslik20=son20haber("baslik")
	hbrresim20=son20haber("resim")
	hbrmetin20=Left(son20haber("metin"),160)
elseif ilkhaber=20 then
	hid_21=son20haber("hid")
	hbrbaslik21=son20haber("baslik")
	hbrresim21=son20haber("resim")
	hbrmetin21=Left(son20haber("metin"),160)
end if
ilkhaber=ilkhaber+1
son20haber.Movenext
Loop
son20haber.close
set son20haber=nothing
%>
<script language="JavaScript">    
function haberlerigoster(maxinuke)
{
    hide('maxinukehaber0a');hide('maxinukehaber0b');hide('maxinukehaber0c');
    hide('maxinukehaber1a');hide('maxinukehaber1b');hide('maxinukehaber1c');
    hide('maxinukehaber2a');hide('maxinukehaber2b');hide('maxinukehaber2c');
    hide('maxinukehaber3a');hide('maxinukehaber3b');hide('maxinukehaber3c');
    hide('maxinukehaber4a');hide('maxinukehaber4b');hide('maxinukehaber4c');
    hide('maxinukehaber5a');hide('maxinukehaber5b');hide('maxinukehaber5c');
    hide('maxinukehaber6a');hide('maxinukehaber6b');hide('maxinukehaber6c');
    hide('maxinukehaber7a');hide('maxinukehaber7b');hide('maxinukehaber7c');
    hide('maxinukehaber8a');hide('maxinukehaber8b');hide('maxinukehaber8c');
    hide('maxinukehaber9a');hide('maxinukehaber9b');hide('maxinukehaber9c');
    hide('maxinukehaber10a');hide('maxinukehaber10b');hide('maxinukehaber10c');
    hide('maxinukehaber11a');hide('maxinukehaber11b');hide('maxinukehaber11c');
    hide('maxinukehaber12a');hide('maxinukehaber12b');hide('maxinukehaber12c');
    hide('maxinukehaber13a');hide('maxinukehaber13b');hide('maxinukehaber13c');
    hide('maxinukehaber14a');hide('maxinukehaber14b');hide('maxinukehaber14c');
    hide('maxinukehaber15a');hide('maxinukehaber15b');hide('maxinukehaber15c');
    hide('maxinukehaber16a');hide('maxinukehaber16b');hide('maxinukehaber16c');
    hide('maxinukehaber17a');hide('maxinukehaber17b');hide('maxinukehaber17c');
    hide('maxinukehaber18a');hide('maxinukehaber18b');hide('maxinukehaber18c');
    hide('maxinukehaber19a');hide('maxinukehaber19b');hide('maxinukehaber19c');
    hide('maxinukehaber20a');hide('maxinukehaber20b');hide('maxinukehaber20c');
	document.getElementById('maxinukehaber'+maxinuke+'a').style.display = 'block';
    document.getElementById('maxinukehaber'+maxinuke+'b').style.display = 'block';
    document.getElementById('maxinukehaber'+maxinuke+'c').style.display = 'block';
}
function hide(maxinuke)
{	document.getElementById(''+maxinuke+'').style.display = 'none';}
</script>
<table width="100%" height="316" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=content_bg%>" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td height="25" colspan="2" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><font size="4"><b>
		<div id="maxinukehaber0b"><%=hbrbaslik1%></div>
		<div id="maxinukehaber1b" style="DISPLAY: none"><%=hbrbaslik2%></div>
		<div id="maxinukehaber2b" style="DISPLAY: none"><%=hbrbaslik3%></div>
		<div id="maxinukehaber3b" style="DISPLAY: none"><%=hbrbaslik4%></div>
		<div id="maxinukehaber4b" style="DISPLAY: none"><%=hbrbaslik5%></div>
		<div id="maxinukehaber5b" style="DISPLAY: none"><%=hbrbaslik6%></div>
		<div id="maxinukehaber6b" style="DISPLAY: none"><%=hbrbaslik7%></div>
		<div id="maxinukehaber7b" style="DISPLAY: none"><%=hbrbaslik8%></div>
		<div id="maxinukehaber8b" style="DISPLAY: none"><%=hbrbaslik9%></div>
		<div id="maxinukehaber9b" style="DISPLAY: none"><%=hbrbaslik10%></div>
		<div id="maxinukehaber10b" style="DISPLAY: none"><%=hbrbaslik11%></div>
		<div id="maxinukehaber11b" style="DISPLAY: none"><%=hbrbaslik12%></div>
		<div id="maxinukehaber12b" style="DISPLAY: none"><%=hbrbaslik13%></div>
		<div id="maxinukehaber13b" style="DISPLAY: none"><%=hbrbaslik14%></div>
		<div id="maxinukehaber14b" style="DISPLAY: none"><%=hbrbaslik15%></div>
		<div id="maxinukehaber15b" style="DISPLAY: none"><%=hbrbaslik16%></div>
		<div id="maxinukehaber16b" style="DISPLAY: none"><%=hbrbaslik17%></div>
		<div id="maxinukehaber17b" style="DISPLAY: none"><%=hbrbaslik18%></div>
		<div id="maxinukehaber18b" style="DISPLAY: none"><%=hbrbaslik19%></div>
		<div id="maxinukehaber19b" style="DISPLAY: none"><%=hbrbaslik20%></div>
		<div id="maxinukehaber20b" style="DISPLAY: none"><%=hbrbaslik21%></div>
		</b></font>
		</td>
	</tr>
	<tr>
		<td width=276 valign="top">
		<div id="maxinukehaber0c" style="width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim1<>"" then%><%=hbrresim1%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber1c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim2<>"" then%><%=hbrresim2%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber2c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim3<>"" then%><%=hbrresim3%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber3c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim4<>"" then%><%=hbrresim4%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber4c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim5<>"" then%><%=hbrresim5%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber5c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim6<>"" then%><%=hbrresim6%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber6c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim7<>"" then%><%=hbrresim7%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber7c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim8<>"" then%><%=hbrresim8%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber8c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim9<>"" then%><%=hbrresim9%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber9c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim10<>"" then%><%=hbrresim10%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber10c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim11<>"" then%><%=hbrresim11%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber11c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim12<>"" then%><%=hbrresim12%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber12c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim13<>"" then%><%=hbrresim13%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber13c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim14<>"" then%><%=hbrresim14%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber14c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim15<>"" then%><%=hbrresim15%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber15c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim16<>"" then%><%=hbrresim16%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber16c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim17<>"" then%><%=hbrresim17%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber17c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim18<>"" then%><%=hbrresim18%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber18c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim19<>"" then%><%=hbrresim19%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber19c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim20<>"" then%><%=hbrresim20%><%else%>yok.png<%end if%>" border=0></div>
		<div id="maxinukehaber20c" style="DISPLAY: none; width:100%; height:15"><img height=205 width=275 src="uploads/image/haber_resim/<%if hbrresim21<>"" then%><%=hbrresim21%><%else%>yok.png<%end if%>" border=0></div>
		<div style="padding-left:3px; padding-top:3px">
		<div id="maxinukehaber0a" style="width:100%; height:15"><%=hbrmetin1%>..</div>
		<div id="maxinukehaber1a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin2%>..</div>
		<div id="maxinukehaber2a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin3%>..</div>
		<div id="maxinukehaber3a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin4%>..</div>
		<div id="maxinukehaber4a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin5%>..</div>
		<div id="maxinukehaber5a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin6%>..</div>
		<div id="maxinukehaber6a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin7%>..</div>
		<div id="maxinukehaber7a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin8%>..</div>
		<div id="maxinukehaber8a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin9%>..</div>
		<div id="maxinukehaber9a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin10%>..</div>
		<div id="maxinukehaber10a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin11%>..</div>
		<div id="maxinukehaber11a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin12%>..</div>
		<div id="maxinukehaber12a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin13%>..</div>
		<div id="maxinukehaber13a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin14%>..</div>
		<div id="maxinukehaber14a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin15%>..</div>
		<div id="maxinukehaber15a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin16%>..</div>
		<div id="maxinukehaber16a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin17%>..</div>
		<div id="maxinukehaber17a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin18%>..</div>
		<div id="maxinukehaber18a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin19%>..</div>
		<div id="maxinukehaber19a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin20%>..</div>
		<div id="maxinukehaber20a" style="DISPLAY: none; width:100%; height:15"><%=hbrmetin21%>..</div>
		</div>
		</td>
		<td valign=top style="padding-left:10px; padding-right:10px">
		<a onmouseover="haberlerigoster(0)" href="news.asp?Action=Read&hid=<%=hid_1%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik1%><br>
		<a onmouseover="haberlerigoster(1)" href="news.asp?Action=Read&hid=<%=hid_2%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik2%><br>
		<a onmouseover="haberlerigoster(2)" href="news.asp?Action=Read&hid=<%=hid_3%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik3%><br>
		<a onmouseover="haberlerigoster(3)" href="news.asp?Action=Read&hid=<%=hid_4%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik4%><br>
		<a onmouseover="haberlerigoster(4)" href="news.asp?Action=Read&hid=<%=hid_5%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik5%><br>
		<a onmouseover="haberlerigoster(5)" href="news.asp?Action=Read&hid=<%=hid_6%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik6%><br>
		<a onmouseover="haberlerigoster(6)" href="news.asp?Action=Read&hid=<%=hid_7%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik7%><br>
		<a onmouseover="haberlerigoster(7)" href="news.asp?Action=Read&hid=<%=hid_8%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik8%><br>
		<a onmouseover="haberlerigoster(8)" href="news.asp?Action=Read&hid=<%=hid_9%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik9%><br>
		<a onmouseover="haberlerigoster(9)" href="news.asp?Action=Read&hid=<%=hid_10%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik10%><br>
		<a onmouseover="haberlerigoster(10)" href="news.asp?Action=Read&hid=<%=hid_11%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik11%><br>
		<a onmouseover="haberlerigoster(11)" href="news.asp?Action=Read&hid=<%=hid_12%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik12%><br>
		<a onmouseover="haberlerigoster(12)" href="news.asp?Action=Read&hid=<%=hid_13%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik13%><br>
		<a onmouseover="haberlerigoster(13)" href="news.asp?Action=Read&hid=<%=hid_14%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik14%><br>
		<a onmouseover="haberlerigoster(14)" href="news.asp?Action=Read&hid=<%=hid_15%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik15%><br>
		<a onmouseover="haberlerigoster(15)" href="news.asp?Action=Read&hid=<%=hid_16%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik16%><br>
		<a onmouseover="haberlerigoster(16)" href="news.asp?Action=Read&hid=<%=hid_17%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik17%><br>
		<a onmouseover="haberlerigoster(17)" href="news.asp?Action=Read&hid=<%=hid_18%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik18%><br>
		<a onmouseover="haberlerigoster(18)" href="news.asp?Action=Read&hid=<%=hid_19%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik19%><br>
		<a onmouseover="haberlerigoster(19)" href="news.asp?Action=Read&hid=<%=hid_20%>"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif" border="0">&nbsp;<%=hbrbaslik20%><br>
		</td>
	</tr>
</table>
<%
End If

If sett("hv") = True Then
Set yirmidenevvel = Server.Createobject("Adodb.Recordset")
yirmidenevvel.open "SELECT * FROM NEWS where onay=true order by hid desc", Connection, 1, 3
yirmidenevvel.AbsolutePage = 3
For m = 1 To sett("hs")
If yirmidenevvel.Eof Then Exit For
	If yirmidenevvel("resim")="" Then
	resmisii="yok.png"
	else
	resmisii=yirmidenevvel("resim")
	End if
%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="100%"  align="center" style="<%=TableBorder%>" class="td_menu">
	<tr height="25">
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" >&nbsp;&nbsp;<B><%=yirmidenevvel("baslik")%></B></td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" align="right"><%=yirmidenevvel("tarih")%>&nbsp;&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" bgcolor="<%=content_bg%>" style="padding-left:3px;padding-right:3px;padding-top:3px;padding-bottom:3px" colspan="2"><IMG SRC="uploads/image/haber_resim/<%=resmisii%>" height="70" BORDER="0" ALT="<%=yirmidenevvel("baslik")%>" align="left"><%=Left(yirmidenevvel("metin"),400)%><%If Len(yirmidenevvel("metin"))>400 Then%>... <A HREF="news.asp?Action=Read&hid=<%=yirmidenevvel("hid")%>"><B>Devamýný Oku</B></A><%End If%></td>
	</tr>
</table>
<%
yirmidenevvel.Movenext
Next
yirmidenevvel.Close
Set yirmidenevvel = Nothing
End if
%>