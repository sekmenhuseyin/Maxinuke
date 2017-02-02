<table width="100%" height="25" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">
  <tr style="font-weight:bold"> 
<%
IF Session("Enter") = "Yes" Then
Set uyebilgi = Connection.Execute("Select * FROM MEMBERS WHERE uye_id = "&Session("Uye_ID")&"")
Set unR = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND READ = False AND SEE = True AND TO = '"&Session("Member")&"'")
Set allmsgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&Session("Member")&"'")
	If unR("count") > sett("msg_limit") Then
	unRd = sett("msg_limit")
	Else
	unRd = unR("count")
	End If
	If allmsgs("count") > sett("msg_limit") Then
	allM = sett("msg_limit")
	Else
	allM = allmsgs("count")
	End If
%>
		<td align="center"><%=Session("Member")%></td>
	    <td><img src="Images/Site/top_menu_blank.gif"></td>
		<td align="center"><a href="Your_Account.asp"><%=ya_topic%></a></td>
	    <td><img src="Images/Site/top_menu_blank.gif"></td>
		<td align="center"><a href="Your_Account.asp?op=Profile">Profilim</a></td>
	    <td><img src="Images/Site/top_menu_blank.gif"></td>
		<td align="center"><a href="Your_Account.asp?op=Friends">Arkadaþlarým</a></td>
	    <td><img src="Images/Site/top_menu_blank.gif"></td>
		<td align="center"><a href="search.asp?action=Member"><%=uyemenu_uyeara%></a></td>
	    <td><img src="Images/Site/top_menu_blank.gif"></td>
		<td align="center"><a href="msgbox.asp?action=main&id=<%=Session("Uye_ID")%>"><%=uyemenu_msgbox%></a> (<%=unRd%>/<%=allM%>)</td>
		<%IF unRd >= 1 And Session("pmvar") = "" then%>
		<script language="JavaScript">
		alert('Mesaj Kutunuzda Okunmamýþ Mesajlarýnýz var.\n\nHesabým Sayfasýndan mesajlarýnýzý okuyabilirsiniz.');
		</script>
		<%
		Session("pmvar") = "ok"
		End IF
		%>
		<script language="JavaScript">
		function goToURL(form)
		{
		var myindex=form.theme.selectedIndex
		if(!myindex=="")
		{
		window.location.href=form.theme.options[myindex].value;
		}
		}
		</script>
		<form action="?op=RegTheme" method="post">
				<td>
		<select name=theme class="form1" onChange=goToURL(this.form);>
		<option value="">Seciniz</option>
		<%
		Set themes = Server.CreateObject("ADODB.RecordSet")
		themes.open "Select * FROM THEMES where t_active=true ORDER BY t_name ASC",Connection,3,3
		Do Until themes.EoF
		IF Session("SiteTheme") = themes("t_dir") Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		%>
		<option value="your_account.asp?op=RegTheme&tema=<%=themes("t_dir")%>" <%=opt%>><%=themes("t_name")%></option>
		<%
		themes.MoveNext
		Loop
		%>
		</select>
		</td>
		</form>
<%If Session("Level") = "1" OR Session("Level") = "2" OR Session("Level") = "3" OR Session("Level") = "4" OR Session("Level") = "5" Then%>
		<td><a href="admin.asp" target="_blank">Yönetim Paneli</a></td>
<%End if%>
		<td><a href="Your_Account.asp?op=Logout"><img src="images/temalar/<%=Session("SiteTheme")%>/logout.gif" border="0"></a></td>
<%
allmsgs.Close
Set allmsgs = Nothing
unR.Close
Set unR = Nothing

ELSE%>
		<SCRIPT LANGUAGE=JAVASCRIPT>
		function giris_kontrol(form) 
		{
		if (form.kuladi.value == "") {alert("<%=error1%>"); return false; }
		if (form.password.value == "") {alert("<%=error2%>"); return false; }
		if (form.guvenlik.value == "") {alert("<%=sc_err1%>"); return false; }
		return true;
		}
		</SCRIPT>
		<form onSubmit="return giris_kontrol(this)"  action="enter.asp" method="post">
		<%
		Randomize
		G1 =  Int(Rnd *10) +1
		G6 =  Int(Rnd *9) +1
		G3 =  Int(Rnd *9) +1
		G4 =  Int(Rnd *9) +1
		G5 =  Int(Rnd *9) +1
		%>
		<input type="hidden" name="gguvenlik" value="<%=G1%><%=G6%><%=G3%><%=G4%><%=G5%>">
		<td>&nbsp;Kullanýcý Adýnýz : <input type="text" name="kuladi" class="form1" size="20"> <%=uye_password%><input type="password" name="password" class="form1" size="20"> Kodunuz : <%=G1%><%=G6%><%=G3%><%=G4%><%=G5%><input type="text" name="guvenlik" class="form1" size="5">&nbsp;&nbsp;Beni Hatýrla<INPUT TYPE="checkbox" NAME="hatirlimmi" class="form1" value="1"> <input type="submit" value="<%=uye_submit%>" class="submit" name="submit" onClick="this.form.submit();this.disabled=true; return true;"></td>
	    <td><img src="Images/Site/top_menu_blank.gif" height="20"></td>
		<td align="center"><a href="uye_islem.asp?action=new"><%=uye_new%></a></td>
	    <td><img src="Images/Site/top_menu_blank.gif" height="20"></td>
		<td align="center"><a href="uye_islem.asp?action=lostpass&step=1"><%=uye_lostpwd%></a></td>
		</form>
<% End IF %>
</tr>
</table>