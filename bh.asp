<%
IF Session("Enter")=""then
	If Request.Cookies(""&sett("site_isim")&"")("Enter") = "Yes" Then
uye_id=Request.Cookies(""&sett("site_isim")&"")("uyeid")
gk=Request.Cookies(""&sett("site_isim")&"")("gk")
set bhkontrol = server.createobject("adodb.recordset")
bhk = "select * from MEMBERS where uye_id="&uye_id&" and gk="&gk&""
bhkontrol.open bhk, Connection, 1, 3
	If sett("onaylama")=True And bhkontrol("onay")=False Then
	Response.Redirect "onay.asp"
	response.end
	else
	password=True
	Session("Enter") = "Yes"
	Session("Uye_ID") = bhkontrol("uye_id")
	Session("Member") = bhkontrol("kul_adi")
	Session("Name") = bhkontrol("isim")
	Session("Email") = bhkontrol("email")
	Session("Level") = bhkontrol("seviye")
	Session("Signature") = bhkontrol("imza")
	Session("Theme") = bhkontrol("u_theme")
	Response.Cookies(""&sett("site_isim")&"_Forum")("LastTime") = ""&Session("son_tarih")&""
	Session.LCID = 1033
	Set up1 = Connection.Execute("UPDATE MEMBERS SET son_tarih=now() WHERE kul_adi='"&Session("Member")&"'")
	Set up2 = Connection.Execute("UPDATE MEMBERS SET login_sayisi="&bhkontrol("login_sayisi")&" + 1 WHERE kul_adi='"&Session("Member")&"'")
	End If
bhkontrol.close
set bhkontrol=nothing
	Else
	End if
End if
%>