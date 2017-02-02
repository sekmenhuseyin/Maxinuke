<%
set aboneleribul = server.createobject("adodb.recordset")
SQL = "Select * from f_abone where kategoriid="&ms("kategoriid")&" order by id asc"
aboneleribul.open SQL,Connection,1,3
for k=1 to aboneleribul.RecordCount
if aboneleribul.eof or aboneleribul.bof then exit for
set mailibul = server.createobject("adodb.recordset")
SQL = "Select * from MEMBERS where gk="&aboneleribul("gk")&""
mailibul.open SQL,Connection,1,3
mailisi=mailibul("email")
html = "<font face=verdana size=2>Merhaba <br><br>"&Now()&" tarihinde "&sett("site_isim")&" Forumlarýndaki abone olduðunuz bir konuya mesaj yazýldý, konuya gitmek içim <A HREF=""#"">BURAYI</A> týklayabilirsiniz. Dilerseniz Hesabým sayfasýndan aboneliklerinizi iptal edebilirsiniz.<BR><BR>Bilgilerinize.</font>"
Set mail = CreateObject("CDONTS.NewMail")
mail.BodyFormat=0
mail.MailFormat=0
mail.Subject=""&sett("site_isim")&" Forum Mesaji Eklendi"
mail.Body=HTML
mail.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
mail.to=mailisi
mail.Importance=1
mail.Send
set HTML = nothing
set mail=nothing
mailibul.close
set mailibul=nothing
aboneleribul.movenext : Next
aboneleribul.close
set aboneleribul=nothing
%>