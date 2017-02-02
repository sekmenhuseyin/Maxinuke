<%Server.ScriptTimeout=900%>
<%Session.Timeout=1300%>
<%Session.CodePage = 1254%>
<!--#include file="database.asp" -->
<!--#include file="Language_File.asp" -->
<!--#include file="CheckLevels.asp" -->
<!--#include file="GetInfo.asp" -->
<!--#include file="maincodes.asp" -->
<!--#include file="themes.asp" -->
<!--#include file="banned_ip.asp" -->
<!--#include file="filter.asp" -->
<!--#include file="no_entry.asp" -->
<!--#include file="functions.asp" -->
<!--//
/* *************************************************************
Software: MaxiNuke version <%=v%>
Info: http://www.maxinuke.com
Copyright: (C) 2008 - 20009 MaxiNuke
http://www.webncs.com - WEBNCS STUDYOLARINDA HAZIRLANMISTIR.
************************************************************* */
//-->
<!--#include file="encryption.asp" -->
<!--#include file="clear_html_codes.asp" -->
<%Session.LCID = sett("s_lcid")%>
<script>
function sayi(deger) {
var yaz = new String();
var numaralar = "0123456789";
var chars = deger.value.split("");
for (i = 0; i < chars.length; i++) {
if (numaralar.indexOf(chars[i]) != -1) yaz += chars[i];
}
if (deger.value != yaz) deger.value = yaz;
}
</script>
<script>
function harf(deger) {
var yaz = new String();
var harfler = "QWERTYUIOPÐÜASDFGHJKLÞÝZCVBNMÖÇqwertyuýopðüasdfghjklþizcvbnmöç ";
var chars = deger.value.split("");
for (i = 0; i < chars.length; i++) {
if (harfler.indexOf(chars[i]) != -1) yaz += chars[i];
}
if (deger.value != yaz) deger.value = yaz;
}
</script>