<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "all" Then
call bloklar
ElseIF act = "New" Then
call new_block
ElseIF act = "NewReg" Then
call new_block_reg
ElseIF act = "Delete" Then
call block_delete
ElseIF act = "Edit" Then
call block_edit
ElseIF act = "Update" Then
call block_update
ElseIf act = "Passive" Then
call passive
ElseIf act = "Active" Then
call active
End If

Sub new_block
%>
<form action="?action=NewReg" method="post">
<table border="0" cellspacing="1" cellpadding="1" align="center" class="td_menu">
	<tr>
		<td align=center><b>-=- Yeni Blok Ekleme Alaný -=-</b></td>
	</tr>
	<tr>
		<td>
		Blok Adý : <input type="text" name="name" class="form1">
		Blok Konumu : 
		<select name="align" size="1" class="form1">
		<option value="LEFT">Sol Blok</option>
		<option value="RIGHT">Sað Blok</option>
		</select>
		Blok Klasöru : <input type="text" name="include" value="null" class="form1">
		Sýra : 
		<select name="queue" size="1" class="form1">
		<% Set block_record = Connection.Execute("Select Count(*) as sayi FROM bloklar")
		For I = 1 TO block_record("sayi")/2+1
		%>
		<option value="<%=I%>"><%=I%></option>
		<%
		Next
		%>
		</select>
		<BR><BR>
		</td>
	</tr>
	<tr>
		<td><textarea name="content" rows="20" class="form1" style="width:100%"></textarea></td>
	</tr>
	<tr>
		<td align="center"><BR><BR><input type="submit" value="BLOKU EKLE" class="submit"></td>
	</tr>
</table>
</form>
<%
End Sub

Sub new_block_reg
name = Request.Form("name")
content = Request.Form("content")
include = Request.Form("include")
align = Request.Form("align")
queue = Request.Form("queue")
actv = "True"

IF name="" or include = "" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlar Doldurunuz</div>"
ELSE
Connection.Execute("INSERT INTO bloklar (b_name, b_content, b_include, b_align, b_active, b_queue) VALUES ('"&name&"', '"&content&"', '"&include&"', '"&align&"', "&actv&", "&queue&")")
Response.Redirect "?action=all"
End IF
End Sub

Sub bloklar
Set rs = Connection.Execute("Select * FROM bloklar")
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="7" align="center"><B>-=- BLOKLAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Blok Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Blok Ýsmi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Konumu</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Sýralamasý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Klasoru</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do While NOT rs.Eof %>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=rs("bid")%></td>
		<td><%=rs("b_name")%></td>
		<td><%=rs("b_align")%></td>
		<td align="center"><%=rs("b_queue")%></td>
		<td><%=rs("b_include")%></td>
		<td><% If rs("b_active") = True Then %>Aktif Yayinda<%Else%><font color=red>Pasif Gosterilmiyor</font><% End If %></td>
		<td align="center"><% If rs("b_active") = True Then %><a href="?action=Passive&id=<%=rs("bid")%>">Pasifleþtir</a><%Else%><a href="?action=Active&id=<%=rs("bid")%>">Aktifleþtir</a><% End If %> - <a href="?action=Edit&id=<%=rs("bid")%>">Düzenle</a> - <a href="?action=Delete&id=<%=rs("bid")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End Sub

Sub passive
blid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE bloklar SET b_active = False WHERE bid = "&blid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub active
blid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE bloklar SET b_active = True WHERE bid = "&blid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub block_edit
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM bloklar WHERE bid = "&Request.QueryString("id")&"",Connection,3,3
%>	
<form action="?action=Update&id=<%=Request.QueryString("id")%>" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu">
	<tr>
		<td align=center><b>-=- Blok Düzenleme Alaný -=-</b></td>
	</tr>
	<tr>
		<td>
		Blok Adý : <input type="text" name="name" value="<%=rs("b_name")%>" class="form1">
		Blok Konumu : 
		<%
		IF rs("b_align") = "LEFT" Then
		opt1 = "selected"
		ElseIF rs("b_align") = "RIGHT" Then
		opt2 = "selected"
		ELSE
		opt1 = "selected"
		End IF
		%>
		<select name="align" size="1" class="form1">
		<option value="LEFT" <%=opt1%>>Sol Blok</option>
		<option value="RIGHT" <%=opt2%>>Sað Blok</option>
		</select>
		Blok Klasöru : <input type="text" name="include" value="<%=rs("b_include")%>" class="form1">
		Sýra : 
		<select name="queue" size="1" class="form1">
		<% Set block_record = Connection.Execute("Select Count(*) as sayi FROM bloklar")
		For I = 1 TO block_record("sayi")+2
		IF I = rs("b_queue") Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<%
		Next
		%>
		</select>
		<BR><BR>
		</td>
	</tr>
	<tr>
		<td><textarea name="content" rows="20" class="form1" style="width:100%"><%=rs("b_content")%></textarea></td>
	</tr>
	<tr>
		<td align="center"><BR><BR><input type="submit" value="BLOKU EKLE" class="submit"></td>
	</tr>
</table>
</table>
</form>
<%
End Sub

Sub block_update
name = Request.Form("name")
content = Request.Form("content")
include = Request.Form("include")
align = Request.Form("align")
queue = Request.Form("queue")

IF name="" or queue="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlar Doldurunuz.</div>"
ELSE
Set upd = Server.CreateObject("ADODB.RecordSet")
upd.open "Select * FROM bloklar WHERE bid = "&Request.QueryString("id")&"",Connection,3,3

upd("b_name") = name
upd("b_content") = content
upd("b_include") = include
upd("b_align") = align
upd("b_queue") = queue
upd.Update
Response.Redirect "?action=all"
End IF
End Sub

Sub block_delete
Connection.Execute("DELETE * FROM bloklar WHERE bid = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

End If
%>          