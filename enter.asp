<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
kuladi = duz(Request.Form("kuladi"))
password = duz(Request.Form("password"))
guvenlik = duz(Request.Form("guvenlik"))
gguvenlik = duz(Request.Form("gguvenlik"))
If not gguvenlik = guvenlik Then
response.write WriteMsg(sc_err2)
Else
set rs = Connection.Execute("SELECT * FROM MEMBERS")
Do Until rs.Eof
IF ucase(kuladi) = ucase(rs("kul_adi")) Then
kul_adi=True
IF MN_PP(password) = rs("sifre") Then 
Session("onay") = rs("onay")
	If sett("onaylama") = True And Session("onay") = False Then
	Response.Redirect "onay.asp"
	response.end
	else
password=True
Session("Enter") = "Yes"
Session("Uye_ID") = rs("uye_id")
Session("Member") = rs("kul_adi")
Session("Name") = rs("isim")
Session("Email") = rs("email")
Session("Level") = rs("seviye")
Session("Signature") = rs("imza")
Session("Theme") = rs("u_theme")
Session("gk") = rs("gk")
	If Request.Form("hatirlimmi") = "1" then
	Response.Cookies(""&sett("site_isim")&"")("Enter") = ""&Session("Enter")&""
	Response.Cookies(""&sett("site_isim")&"")("gk") = ""&rs("gk")&""
	Response.Cookies(""&sett("site_isim")&"")("uyeid") = ""&Session("Uye_ID")&""
	Response.Cookies(""&sett("site_isim")&"").Expires = Date() + 90
	End if
Response.Cookies(""&sett("site_isim")&"_Forum")("LastTime") = ""&Session("son_tarih")&""
Session.LCID = 1033
Set up1 = Connection.Execute("UPDATE MEMBERS SET son_tarih=now() WHERE kul_adi='"&Session("Member")&"'")
Set up2 = Connection.Execute("UPDATE MEMBERS SET login_sayisi="&rs("login_sayisi")&" + 1 WHERE kul_adi='"&Session("Member")&"'")
Set delete_user = Connection.Execute("DELETE * FROM GUESTS WHERE ip = '"&Request.ServerVariables("REMOTE_ADDR")&"'")
'	Response.Redirect "Your_Account.asp"
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	End if
End If
End If
rs.MoveNext
Loop

If kul_adi <> True then
Response.Write WriteMsg(error7)
ElseIf password <> True then
Response.Write WriteMsg(error8)
End If

rs.Close
Set rs = Nothing
End IF
call ORTA2
call ALT
%>