<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												Menu Blok Uygulamasi V 1.3											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
Set cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_queue")
Do Until cats.EoF
Response.Write "<b>» " & cats("mc_title") & "</b><br>"
Set module = Connection.Execute("SELECT * FROM moduller WHERE mactive = True AND mcat = "&cats("mc_id")&" ORDER BY mname ASC")
Set links = Connection.Execute("Select * FROM MENU_LINKS WHERE m_status = True AND m_cat = "&cats("mc_id")&"")
Set pages = Connection.Execute("SELECT * FROM PAGES WHERE onay=true and p_cat = "&cats("mc_id")&" ORDER BY p_title ASC")
Do Until links.EoF
acilis=links("m_style")
IF acilis = "normal" Then
strStyle=""
ELSE
strStyle=acilis
End IF
Response.Write "<img src=""IMAGES/Site/menu_select.gif"" align=""absmiddle""> <a href=""http://www."&links("m_url")&""" target="&strStyle&">" & links("m_name") & "</a><br>"
links.MoveNext
Loop
Do Until module.Eof
name = module("mname")
module_id = module("module_id")
%>
<img src="IMAGES/Site/menu_select.gif" align="absmiddle"> <a href="moduller.asp?mi=<%=module_id%>"><%=name%></a><br>
<%
module.MoveNext
Loop
Do Until pages.Eof
Response.Write "<img src=""IMAGES/Site/menu_select.gif"" align=""absmiddle""> <a href=""sayfa.asp?id="&pages("p_id")&""">"&pages("p_title")&"</a><br>"
pages.MoveNext
Loop
cats.MoveNext
Loop
Set pages = Nothing
Set links = Nothing
Set module = Nothing
cats.Close
Set cats = Nothing
%>