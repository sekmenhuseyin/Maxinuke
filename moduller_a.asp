<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
klasor = Server.MapPath("moduller/")
IF Request.QueryString("action") = "EditModule" Then
mdl_id = Request.QueryString("id")
Set mdlW = Server.CreateObject("ADODB.RecordSet")
mdlSQL = "Select * FROM moduller WHERE module_id = "&mdl_id&""
mdlW.open mdlSQL,Connection,1,3
Function KlasorModule(klasorlist)
  Dim fso, f, f1, f2, fc, ss
  Set fso = CreateObject("Scripting.FileSystemObject" )
  Set f = fso.GetFolder(klasorlist)
  Set fc = f.SubFolders
  For Each f in fc
   IF f.name = mdlW("mdir") Then
	opt = "selected"
   ELSE
	opt = ""
   End IF
    ss = ss & "<option value="""&f.name&""" "&opt&">"
    ss = ss & ""&f.name&""
    ss = ss & "</option>"
    KlasorModule = ss
  Next
End Function
ELSE
Function KlasorListele(klasorlist)
  Dim fso, f, f1, f2, fc, s
  Set fso = CreateObject("Scripting.FileSystemObject" )
  Set f = fso.GetFolder(klasorlist)
  Set fc = f.SubFolders
  For Each f in fc
    s = s & "<option value="""&f.name&""">"
    s = s & "" & f.name & ""
    s = s & "</option>"
    KlasorListele = s
  Next
End Function
End IF
act = Request.QueryString("action")
If act = "all" Then
call all
ElseIF act = "New" Then
call new_module
ElseIf act = "Register" Then
call reg
ElseIf act = "Active" Then
call active
ElseIf act = "Passive" Then
call passive
ElseIf act = "DeleteModule" Then
call delete
ElseIf act = "EditModule" Then
call edit
ElseIf act = "UpdateModule" Then
call update
End If

Sub all
Set mds = Connection.Execute("Select * FROM moduller")
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="6" align="center"><B>-=- MODULLER -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Modül Ýsmi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Sýrasý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Görebilenler</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do While NOT mds.Eof
Set cat = Connection.Execute("Select * FROM MENU_CATS WHERE mc_id = "&mds("mcat")&"")
%>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td><%=mds("mname")%></td>
		<td align="center"><%=cat("mc_title")%></td>
		<td align="center">-</td>
		<td align="center"><% If mds("mactive") = True Then %>Aktif&nbsp;<%Else%><font color=red>Pasif</font><% End If %></td>
		<td align="center"><% If mds("mems") = True Then %>Sadece Üyeler<%Else%>Herkes<% End If %></td>
		<td align="center"><% If mds("mactive") = True Then %><a href="?action=Passive&id=<%=mds("module_id")%>">Pasifleþtir</a><%Else%><a href="?action=Active&id=<%=mds("module_id")%>"><font color="red"><B>Aktifleþtir</B></font></a><% End If %> - <a href="?action=EditModule&id=<%=mds("module_id")%>">Düzenle</a> - <a href="?action=DeleteModule&id=<%=mds("module_id")%>">Sil</a></td>
	</tr>
<%
mds.MoveNext
Loop
%>
</table>
<%
End Sub
'####################################################################################################################################
Sub new_module
Set mcats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC")
%>
<form action="?action=Register" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="50%" align="center" class=td_menu>
	<tr>
		<td>Modülünüzün Görünecek Adý : </td>
		<td><input type="text" name="mname" class="form1" size="20"></td>
	</tr>
	<tr>
		<td>Modülünüzün Dosyasý : </td>
		<td>
		<select name="mdir" size="1" class="form1">
		<%=KlasorListele(klasor)%>
		</select>
		</td>
	</tr>
	<tr>
		<td>Kategorisi : </td>
		<td>
		<select name="mcat" size="1" class="form1">
		<% Do Until mcats.EoF %>
		<option value="<%=mcats("mc_id")%>"><%=mcats("mc_title")%></option>
		<%
		mcats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td>Sadece Üyeler Görebilsin : </td>
		<td><input type="checkbox" name="mems"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="submit" value="Modülümü Ekle" class="submit" style="width:100%"></td>
	</tr>
</table>
</form>
<%
End Sub
'####################################################################################################################################
Sub reg
name = duz(Request.Form("mname"))
dir = duz(Request.Form("mdir"))
cat = Request.Form("mcat")

If Request.Form("mems") = "on" Then
mem = True
Else
mem = False
End If

If name="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Ýsmini Yazýnýz...</font></center>"
ElseIf dir="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Dosya Adresini Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM moduller"
ent.open eSQL,Connection,1,3

ent.AddNew
ent("mname") = name
ent("mdir") = dir
ent("mems") = mem
ent("mcat") = cat
ent("mactive") = true
ent.Update
Response.redirect "?action=all"
END IF
End Sub
'####################################################################################################################################
Sub active
mdlid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE moduller SET mactive = True WHERE module_id = "&mdlid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
'####################################################################################################################################

Sub passive
mdlid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE moduller SET mactive = False WHERE module_id = "&mdlid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
'####################################################################################################################################
Sub delete
mdlid = Request.QueryString("id")
Set dm = Server.CreateObject("ADODB.RecordSet")
dmSQL = "DELETE * FROM moduller WHERE module_id = "&mdlid&""
dm.open dmSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
'####################################################################################################################################
Sub edit
mdid = Request.QueryString("id")
Set mdl = Connection.Execute("Select * FROM moduller WHERE module_id = "&mdid&"")
Set cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC")
%>
<form action="?action=UpdateModule&id=<%=mdid%>" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="50%" align="center" class=td_menu>
	<tr>
		<td>Modül Ýsmi : </td>
		<td><input type="text" name="mname" class="form1" size="20" value="<%=mdl("mname")%>"></td>
	</tr>
	<tr>
		<td>Modül Dosyasý : </td>
		<td>
		<select name="mdir" size="1" class="form1">
		<%=KlasorModule(klasor)%>
		</select>
		</td>
	</tr>
	<tr>
		<td>Kategori : </td>
		<td>
		<select name="mcat" size="1" class="form1">
		<%
		Do Until cats.EoF
		IF cats("mc_id") = mdl("mcat") Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		%>
		<option value="<%=cats("mc_id")%>" <%=opt%>><%=cats("mc_title")%></option>
		<%
		cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td>Sadece Üyeler Görebilsin : </td>
		<td>
		<%
		If mdl("mems") = True Then
		modS = "checked"
		ElseIf mdl("mems") = False Then
		modS = ""
		End If
		%>
		<input type="checkbox" name="mems" <%=modS%>></td>
	</tr>
	<tr>
		<td colspan="2" ><input type="submit" value="Güncelle" class="submit" style="width:100%"></td>
	</tr>
</table>
</form>
<%
End Sub
'####################################################################################################################################
Sub update
mdid = Request.QueryString("id")
name = duz(Request.Form("mname"))
dir = duz(Request.Form("mdir"))
mcat = duz(Request.Form("mcat"))

If Request.Form("mems") = "on" Then
mem = True
Else
mem = False
End If

If name="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Ýsmini Yazýnýz...</font></center>"
ElseIf dir="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Dosya Adresini Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM moduller WHERE module_id = "&mdid&""
ent.open eSQL,Connection,1,3
ent("mname") = name
ent("mdir") = dir
ent("mems") = mem
ent("mcat") = mcat
ent.Update
Response.redirect "?action=all"
END IF
End Sub
End If
%>