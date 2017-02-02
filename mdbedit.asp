<%@LANGUAGE="VBSCRIPT" codepage=1254%>
<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "" Or Not Session("Level") = "1" Then
response.redirect "default.asp"
else
On Error Resume Next
Session.Timeout = 1000
Response.Buffer = True
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim wordsactive, IE, check, head13, header13, td1c, td4c, td5c, deep, ft2c	,buadres, basla, busayfa, speciapenouserldb, penouser
Dim hata(2,2)
penouser = "admin"
penopass = "admin"
wordsactive = True
'********************************************************************************

'THR	||	Internet Explorer'ı kontrol etsin mi? İptal ederseniz tabloya tıklamalar iptal olur / False yazın
checkIE = True 'False
	if (checkIE) then
		if Instr(Request.ServerVariables("HTTP_USER_AGENT"), "IE") => 1 Then IE = True
	End if
'********************************************************************************

'// Eğer specialdb ve specialdsnless ı boş bırakırsınız db seçme ekranına ulaşırsınız
'THR	||	You can set a db for a customer
'If you leave blank specialdb, specialdsnless Then you will access Db Selection menu when you run peno (Supoort Remember DB)
specialdb = "../db/"&veritabani_adi&".mdb" '// Ex : demo.mdb
specialdsnless = "../db/"&veritabani_adi&".mdb" 'Write your mdb name with full path EXM : "peno.mdb" or "../peno.mdb"
specialodbc = "" 'Write your odbc name

'********************************************************************************

'THR	||	Doğru Kelimesi / Acc 97 vs. de dillere göre sorun çıkartırsa bunu deneyin...
Dim dkelimesi, ykelimesi
dkelimesi = "True" ' Doğru
ykelimesi = "False" ' Yanlış
'********************************************************************************

'THR	||	Internet Explorer mı ?
'THR	||	Kod değişkenleri
'THR	||	heading13 = Sayfa Başlık Metinleri
'THR	||	header13 = Headerlar
'THR	||	td1c = İlk sütun rengi
'THR	||	td3c = alt | üst sütunlar
'THR	||	td4c = arka plan
'THR	||	td5c = orta alanlar
'THR	||	ft2c = Başlık İşaretleri
td1c = "#333333"
td3c = "#E4E4E4"
td4c = "#999999"
td5c = "#F7F7F7"
ft2c = "#FF6600"
header13 = "<html><head><title>peno DB Editor " & penover & "</title><meta http-equiv=Content-Type content=text/html; charset=ISO-8859-9>"
head13 = "<style type=text/css><!--" & "input {  width: 350px; border: #333333; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; font-family: Geneva, Verdana; font-size: 11px}"& vbNewLine & "td {  font-family: Geneva, Verdana; font-size: 11px}"& vbNewLine &  "textarea {  height: 200px; width: 350px; border: #333333; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; font-family: Geneva, Verdana; font-size: 11px}"& vbNewLine &  ".buton { BACKGROUND-COLOR: #f60; BORDER-BOTTOM-COLOR: #FFB17D; BORDER-BOTTOM-STYLE: outset; BORDER-LEFT-COLOR: #FFB17D; BORDER-LEFT-STYLE: outset; BORDER-RIGHT-COLOR: #FFB17D; BORDER-RIGHT-STYLE: outset; BORDER-TOP-COLOR: #FFB17D; BORDER-TOP-STYLE: outset; COLOR: #fff; FONT-WEIGHT: bold; width:80 px}"& vbNewLine &  ".x1 {  font-family: Geneva, Verdana; font-size: 12px; font-weight:bold}"& vbNewLine &  "a:link {  text-decoration: none}"& vbNewLine &  "a:hover { text-decoration: none ; color: #FF0000}"& vbNewLine &  "a:visited { text-decoration: none }"& vbNewLine &  "select {  width: 150px; font-family: Geneva, Verdana; font-size: 12px; color:#FFFFFF; background-color:#000000; BACKGROUND-COLOR: #f60; BORDER-BOTTOM-COLOR: #FFB17D; BORDER-BOTTOM-STYLE: outset; BORDER-LEFT-COLOR: #FFB17D; BORDER-LEFT-STYLE: outset; BORDER-RIGHT-COLOR: #FFB17D; BORDER-RIGHT-STYLE: outset; BORDER-TOP-COLOR: #FFB17D; BORDER-TOP-STYLE: outset; COLOR: #fff; FONT-WEIGHT: bold;}"& vbNewLine &  ".P13 {  font-family: Geneva, Verdana; font-size: 14px; font-weight: bold; color: #FFFFFF}"& vbNewLine &  ".kayitlar {  width: 90px; font-family: Geneva, Verdana; font-size: 11px; color:#FFFFFF; background-color:#000000; BACKGROUND-COLOR: #f60; BORDER-BOTTOM-COLOR: #FFB17D; BORDER-BOTTOM-STYLE: outset; BORDER-LEFT-COLOR: #FFB17D; BORDER-LEFT-STYLE: outset; BORDER-RIGHT-COLOR: #FFB17D; BORDER-RIGHT-STYLE: outset; BORDER-TOP-COLOR: #FFB17D; BORDER-TOP-STYLE: outset; COLOR: #fff; FONT-WEIGHT: bold;}"& vbNewLine &  "a.admin:link {  text-decoration: none ; background-color:#00FF00}"& vbNewLine &  "a.admin:visited {  text-decoration: none ; background-color:#00FF00}"& vbNewLine &  ".deactive { color:#999999}"& vbNewLine &  ".clean {  width: 20px; border: none; border-style: none; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 0px}" & vbNewLine & ".smallput { height: 13px; width: 13px; border-style: solid; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 0px}" & vbNewLine & "--></style></head>"

If Request.Querystring("olay") = "update" Then	head13 = head13 & "<body leftmargin=2 topmargin=2 marginwidth=0 marginheight=0 link=#FF6600 vlink=#FF6600 onLoad=""focuser()"">" Else 	head13 = head13 & "<body leftmargin=2 topmargin=2 marginwidth=0 marginheight=0 link=#FF6600 vlink=#FF6600>"
deep = "</body></html>"

'***	||	END OF EDITABLE REIGONS	||
'********************************************************************************

penover = "1.7"
penono = "Hayır"
penoyes = "Evet"
penook = "Tamam"
penofirst = "İlk Sayfa"
penoback = "Geri"
penonext = "İleri"
penolast = "Son sayfa"
penomain = "Ana Sayfa"
penonewr = "Yeni Kayıt Ekle"
penodelr = "Kaydı Sil"
penoeditr = "Kaydı Değiştir"
penoemail = "E-mail"
penoblankch = " b "
penoor = " veya "
penoand = " ve "
penodel = "Sil"
penoback = "Geri Dön"

k1 = "Veri tabanını Düzenle / Onar"
k2 = "Veri tabanını Düzenle / Onar"
k3 = "adresindeki Veri tabanınıza Düzenle / Onar (compact) uygulanılacak."
k4 = "Veri tabanınız "

k6 = "Bu işlemin süresi veri tabanının boyutuna göre değişiklik gösterebilir. "
k7 = "Bu işlem sırasında veritabanına girilen yeni kayıtlar kaybolacaktır."
k8 = "Bu işlemi yapmak istediğinize emin misiniz ?"
k9 = ".mdb Veri tabanını Düzenle / Onar tamamlandı"
k10 = "Veri tabanınız başarıyla Düzenlendi."
k11 = "saniye"
k12 = "içerisinde işlem tamamlandı."
k13 = "boyutunda olan veritabanı artık"
k14 = "Tam olarak "
k15 = "dosya boyutu küçüldü"
k16 = "Yeni Kayıt Ekleme"
k17 = "alanı boş olamaz"
k18 = " alanı sayı olmak zorunda"
k19 = "Kayıt eklerken bir kaç sorun oluştu..."
k20 = "Boş olabilir"
k21 = "Boş olamaz"
k22 = "OLE Nesnesi"
k23 = "Veritabanı Seç !"
k24 = "Veritabanı adresini girin..."
k25 = "gibi"
k26 = "Eğer veritabanınız şifreli ise şifrenizi girin."
k27 = "Bağlantı tipinizi seçin. DSN ile veritabanınız sisteme tanıltılmış ise DSN. Eğer sisteme tanıtılmadan direk bağlantı kuruluyorsa DSNLess/OLEDB seçeneğini seçin"
k28 = "Klasörler arasında veritabanınızı bularakta veritabanınıza bağlanabilirsiniz"
k29 = "Üst Klasör"
k30 = "Klasörler"
k31 = "Bu klasördeki veritabanlarını listele..."
k32 = "DSNLess Bağlan..."
k34 = "numaralı kaydı silmek istediğinizden emin misiniz ?"
k35 = "metin"
k36 = "Eklemek istediğiniz resmin adresi : "
k37 = "Tablo Seç"
k39 = "Kaydı Değiştir"
k40 = "numaralı kayıt"
k41 = "tablosunda bulunamadı. "
k42 = "Linki giriniz : "
k43 = "Metni Giriniz : "
k44 = "E-mail Linki Ekle "
k45 = "E-mail Adresiniz giriniz : "
k47 = "Resim Ekle"
k48 = "Resim"
k49 = "Koyu Yazı Yaz"
k50 = "Kalın Gösterilecek Metni Giriniz :"
k51 = "İtalik Yazı Yaz"
k52 = "İtalik Gösterilecek Metni Giriniz :"
k53 = "Altı Çizili Yazı Yaz"
k54 = "Altı Çizili Gösterilecek Metni Giriniz :"
k55 = "Yatay Çizgi Ekle"
k56 = "Altı Çizili Yazı Yaz"
k57 = "Aradığınız kriterlere göre kayıt bulunamadı !. "
k58 = "Kayıtlara Geri Dön"
k59 = "Bu Tabloda Hiç Kayıt Yok !"
k60 = "toplam"
k61 = "kayıt var"
k62 = "Kayıtları Yenile"
k63 = "Arama İptal"
k64 = "Listele"
k65 = "Kayıt No (pk) :"
k66 = "Azdan Çoğa doğru sırala"
k67 = "Çoktan Aza doğru sırala"
k68 = "değiştir"
k69 = "sil"
k70 = "Arama Yapın"
k71 = "Hata no :"
k72 = "Tekrar deneyin !"
k73 = " Kayıt"
k74 = "Link Ekle"
k75 = " Numaralı Kaydı Güncelleme "
k76 = " Tablosundaki Kayıtlar"

'// Beni Hatırla
If Request.Cookies("penologin")("user") = penouser AND Request.Cookies("penologin")("pass") = penopass Then Session("penouser") = penouser

' Bu sayfa
'********************************************
	burasi = Request.Servervariables("PATH_INFO")
	basla = InstrRev(burasi,"/",-1,1)
	busayfa = Mid(burasi,basla+1)

' Veritabanı, tablo ve primary key seçimleri
'********************************************
	olay = Request.Querystring("olay")
	kayitno = Request.Querystring("no")
	sayfa   = Request.Querystring("sayfa")
	If sayfa = "" Then sayfa = 1
	sayfa = CLng(sayfa)

	If Request.Querystring("x") = "db" Then Call dbremove

	if Session("db13") = "" Then' Veritabanı seçimi
		If Request.Cookies("db") <> "" Then
			Session("db13") = Request.Cookies("db")
			db13 = Request.Cookies("db")
			Response.Redirect busayfa
		End if
'		Response.End
		Call dbselector()
	Else
	db13 = Session("db13")
			if  Session("td13") = "" AND Request.Querystring("olay") <> "tdselect" then ' Tablo seçimi
				Response.Redirect(busayfa & "?olay=tdselect")
			Else
				td13 = Session("td13")
				pk13 = Session("pk13")
			End if
	End if

' Veritabanına bağlantı stringi
'********************************************
	Dim conn13
	Set conn13 = Server.CreateObject("ADODB.Connection")
		conn13.ConnectionString = db13
		conn13.Open
		hata(0,0) = err.description
		hata(0,1) = err.number

' Gönderme İşlemleri // Olaylar
'********************************************
	Set Session("dbname13") = conn13
	Select Case olay
		Case "list"							'Kayıt Listesi
			Call listele()
		Case "update"						'Güncelleme
			Call guncelle(kayitno)
		Case "updateOk"					'Güncelleme OK !
			Call guncelleOk(kayitno)
		Case "insert"						'Yeni kayıt ekle
			Call Ekle()
		Case "insertOk"					'Kayıt Ekle Ok !
			Call EkleOk()
		Case "delete"						'Silme onayı
			Call Sil(kayitno)
		Case "deleteOk"					'Sil OK !
			Call SilOk(kayitno)
		Case "schema"
			Call schema13()				' Veritabanı scheması
		Case "dbselect"
			Call dbselector()				' Veritabanı seçimi
		Case "tdselect"
			Call tdselector()				' Tablo seçimi
		Case "restart"
			Call restart()					' Arama ve Listeleme özelliklerini Kaldır
		Case "compact"
			Call compact()					' Veri Tabanını Düzenle / Onar
		Case "logout"
			Call logout()					' Logout / Çıkış
		Case Else							
			Call listele()						'Kayıt Listesi
	End Select

'********************************************
Function logout() ' Logout / Çıkış
'********************************************
Session("penouser") = ""
Session.Contents.Remove("penouser")
Session.Abandon
Response.Cookies("penologin").Expires = Date - 1000

Response.Redirect busayfa
End function

'********************************************
Function dict(kelime) ' Mevcut kelimeleri güzelleri ile değiştir
'********************************************
allwords = Split(words,"|",-1,1)
	For i = 0 to Ubound(allwords)
	Next
	If kelimeval <> "" Then 	Response.Write kelimeval Else Response.Write kelime
End Function


'********************************************
Function cont() ' Veri giriş yapılırken gereken kontrol sistemi
'********************************************
Response.Write "<script language=""JavaScript"">"
Response.Write "<!--"& vbNewLine
Response.Write "function FDK_Validate(FormName, stopOnFailure, AutoSubmit, ErrorHeader)"& vbNewLine
Response.Write "{"& vbNewLine
Response.Write " var theFormName = FormName;"& vbNewLine
Response.Write " var theElementName = """";"& vbNewLine
Response.Write " if (theFormName.indexOf(""."")>=0)  "& vbNewLine
Response.Write " {"& vbNewLine
Response.Write "   theElementName = theFormName.substring(theFormName.indexOf(""."")+1)"& vbNewLine
Response.Write "   theFormName = theFormName.substring(0,theFormName.indexOf("".""))"& vbNewLine
Response.Write " }"& vbNewLine
Response.Write " var ValidationCheck = eval(""document.""+theFormName+"".ValidateForm"")"& vbNewLine
Response.Write " if (ValidationCheck)  "& vbNewLine
Response.Write " {"& vbNewLine
Response.Write "  var theNameArray = eval(theFormName+""NameArray"")"& vbNewLine
Response.Write "  var theValidationArray = eval(theFormName+""ValidationArray"")"& vbNewLine
Response.Write "  var theFocusArray = eval(theFormName+""FocusArray"")"& vbNewLine
Response.Write "  var ErrorMsg = """";"& vbNewLine
Response.Write "  var FocusSet = false;"& vbNewLine
Response.Write "  var i"& vbNewLine
Response.Write "  var msg"& vbNewLine
Response.Write "    "& vbNewLine
Response.Write " "& vbNewLine
Response.Write "        // Go through the Validate Array that may or may not exist"& vbNewLine
Response.Write "        // and call the Validate function for all elements that have one."& vbNewLine
Response.Write "  if (String(theNameArray)!=""undefined"")"& vbNewLine
Response.Write "  {"& vbNewLine
Response.Write "   for (i = 0; i < theNameArray.length; i ++)"& vbNewLine
Response.Write "   {"& vbNewLine
Response.Write "    msg="""";"& vbNewLine
Response.Write "    if (theNameArray[i].name == theElementName || theElementName == """")"& vbNewLine
Response.Write "    {"& vbNewLine
Response.Write "      msg = eval(theValidationArray[i]);"& vbNewLine
Response.Write "    }"& vbNewLine
Response.Write "    if (msg != """")"& vbNewLine
Response.Write "    {"& vbNewLine
Response.Write "     ErrorMsg += ""\n""+msg;                   "& vbNewLine
Response.Write "     if (stopOnFailure == ""1"") "& vbNewLine
Response.Write "     {"& vbNewLine
Response.Write "       if (theFocusArray[i] && !FocusSet)  "& vbNewLine
Response.Write "      {"& vbNewLine
Response.Write "       FocusSet=true;"& vbNewLine
Response.Write "       theNameArray[i].focus();"& vbNewLine
Response.Write "      }"& vbNewLine
Response.Write "      alert(ErrorHeader+ErrorMsg);"& vbNewLine
Response.Write "      document.MM_returnValue = false; "& vbNewLine
Response.Write "      break;"& vbNewLine
Response.Write "     }"& vbNewLine
Response.Write "     else  "& vbNewLine
Response.Write "     {"& vbNewLine
Response.Write "      if (theFocusArray[i] && !FocusSet)  "& vbNewLine
Response.Write "      {"& vbNewLine
Response.Write "       FocusSet=true;"& vbNewLine
Response.Write "       theNameArray[i].focus();"& vbNewLine
Response.Write "      }"& vbNewLine
Response.Write "     }"& vbNewLine
Response.Write "    }"& vbNewLine
Response.Write "   }"& vbNewLine
Response.Write "  }"& vbNewLine
Response.Write "  if (ErrorMsg!="""" && stopOnFailure != ""1"") "& vbNewLine
Response.Write "  {"& vbNewLine
Response.Write "   alert(ErrorHeader+ErrorMsg);"& vbNewLine
Response.Write "  }"& vbNewLine
Response.Write "  document.MM_returnValue = (ErrorMsg==""""); "& vbNewLine
Response.Write "  if (document.MM_returnValue && AutoSubmit)  "& vbNewLine
Response.Write "  {"& vbNewLine
Response.Write "   eval(""document.""+FormName+"".submit()"")"& vbNewLine
Response.Write "  }"& vbNewLine
Response.Write " }"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function FDK_StripChars(theFilter,theString)"& vbNewLine
Response.Write "{"& vbNewLine
Response.Write "	var strOut,i,curChar"& vbNewLine
Response.Write vbNewLine 
Response.Write "	strOut = """""& vbNewLine
Response.Write "	for (i=0;i < theString.length; i++)"& vbNewLine
Response.Write "	{		"& vbNewLine
Response.Write "		curChar = theString.charAt(i)"& vbNewLine
Response.Write "		if (theFilter.indexOf(curChar) < 0)	// if it's not in the filter, send it thru"& vbNewLine
Response.Write "			strOut += curChar		"& vbNewLine
Response.Write "	}	"& vbNewLine
Response.Write "	return strOut"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function FDK_AddToValidateArray(FormName,FormElement,Validation,SetFocus)"& vbNewLine
Response.Write "{"& vbNewLine
Response.Write "    var TheRoot=eval(""document.""+FormName);"& vbNewLine
Response.Write " "& vbNewLine
Response.Write "    if (!TheRoot.ValidateForm) "& vbNewLine
Response.Write "    {"& vbNewLine
Response.Write "        TheRoot.ValidateForm = true;"& vbNewLine
Response.Write "        eval(FormName+""NameArray = new Array()"")"& vbNewLine
Response.Write "        eval(FormName+""ValidationArray = new Array()"")"& vbNewLine
Response.Write "        eval(FormName+""FocusArray = new Array()"")"& vbNewLine
Response.Write "    }"& vbNewLine
Response.Write "    var ArrayIndex = eval(FormName+""NameArray.length"");"& vbNewLine
Response.Write "    eval(FormName+""NameArray[ArrayIndex] = FormElement"");"& vbNewLine
Response.Write "    eval(FormName+""ValidationArray[ArrayIndex] = Validation"");"& vbNewLine
Response.Write "    eval(FormName+""FocusArray[ArrayIndex] = SetFocus"");"& vbNewLine
Response.Write " "& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function ValidateInteger(FormElement,Required,Minimum,Maximum,ErrorMsg)"& vbNewLine
Response.Write "{"& vbNewLine
Response.Write "	var theString = FormElement.value;"& vbNewLine
Response.Write "	theString = FDK_StripChars("" "",theString);"& vbNewLine
Response.Write "	var min = Minimum;"& vbNewLine
Response.Write "	var max = Maximum;"& vbNewLine
Response.Write "	var pean = ErrorMsg;"& vbNewLine
Response.Write vbNewLine 
Response.Write "	if (theString.length == 0)	"& vbNewLine
Response.Write "	{"& vbNewLine
Response.Write "		if (!Required) return """"		"& vbNewLine
Response.Write "		else return pean;"& vbNewLine
Response.Write "	}"& vbNewLine
Response.Write vbNewLine 
Response.Write "	// remove leading zeros (zeros are only leading if there is more than one char)"& vbNewLine
Response.Write "	while (theString.length > 1 && theString.substring(0,1) == ""0"")"& vbNewLine
Response.Write "	{"& vbNewLine
Response.Write "		theString = theString.substring(1, theString.length);"& vbNewLine
Response.Write "	}"& vbNewLine
Response.Write vbNewLine 
Response.Write "	var val = parseInt(theString);"& vbNewLine
Response.Write "	if (isNaN(val)) return pean;"& vbNewLine
Response.Write "	"& vbNewLine
Response.Write "	// check for non-digits (and minus sign). Do this by converting number"& vbNewLine
Response.Write "	// back to a string and comparing it to original string."& vbNewLine
Response.Write "	if (val.toString() != theString) return pean;"& vbNewLine
Response.Write "	"& vbNewLine
Response.Write "	if (min < max)"& vbNewLine
Response.Write "	{"& vbNewLine
Response.Write "		if ((val < min) || (val > max))"& vbNewLine
Response.Write "		{"& vbNewLine
Response.Write "			return ErrorMsg;"& vbNewLine
Response.Write "		}"& vbNewLine
Response.Write "	}"& vbNewLine
Response.Write "	   "& vbNewLine
Response.Write "	// reset the entered string after removal of spaces and leading zeros."& vbNewLine
Response.Write "	FormElement.value=theString;"& vbNewLine
Response.Write "	return """";"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function FDK_AddIntegerValidation(FormName,FormElementName,Required,Minimum,Maximum,SetFocus,ErrorMsg)  {"& vbNewLine
Response.Write "  var ValString = ""ValidateInteger(""+FormElementName+"",""+Required+"",""+Minimum+"",""+Maximum+"",""+ErrorMsg+"")"""& vbNewLine
Response.Write "  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function FDK_ValidateNonBlank(FormElement,ErrorMsg)"& vbNewLine
Response.Write "{"& vbNewLine
Response.Write "  var msg =ErrorMsg;"& vbNewLine
Response.Write "  var val = FormElement.value;  "& vbNewLine
Response.Write vbNewLine 
Response.Write "  if (!(FDK_StripChars("" \n\t\r"",val).length == 0))"& vbNewLine
Response.Write "  {"& vbNewLine
Response.Write "     msg="""";"& vbNewLine
Response.Write "  }"& vbNewLine
Response.Write vbNewLine 
Response.Write "  return msg;"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function FDK_AddNonBlankValidation(FormName,FormElementName,SetFocus,ErrorMsg)  {"& vbNewLine
Response.Write "  var ValString = ""FDK_ValidateNonBlank(""+FormElementName+"",""+ErrorMsg+"")"""& vbNewLine
Response.Write "  FDK_AddToValidateArray(FormName,eval(FormElementName),ValString,SetFocus)"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write "//-->"
Response.Write "</script>"
End Function

'********************************************
Function restart() 'Arama ve Listeleme özelliklerini Kaldır
'********************************************
	Session("aramasql") = ""
	Call listele()
End Function

'********************************************
Function Ekle() 'Yeni Kayıt Ekleme Formu
'********************************************
	Dim formtxt
	heading13 = "<span class=x1>" & penonewr & "</span>"
	Set kayit13 = Server.CreateObject("ADODB.Recordset")
	kayit13.Open td13, conn13, 1, 2
	'*Field Sayısı
	toplamf = kayit13.Fields.Count - 1
	Response.Write header13
	Call cont()
	Response.Write head13
	Call list13(k16)
	Response.Write 	 "<form name=""form13"" action=""" & busayfa & "?sayfa=" & sayfa & "&olay=insertOk"" method=""POST"" onSubmit=""" 
		For i = 0 To toplamf
		Response.Write i & "<br>"
			fname = kayit13(i).Name
'			fdeger = kayit13(i).Value
			ftip = kayit13(i).Type
			fatt = kayit13(i).Attributes
			fsize = kayit13(i).definedsize
			If fname <> pk13 Then
				' Boş kabul edilmiyorsa
				If (fatt AND &H00000020) <> &H00000020 Then Response.Write "FDK_AddNonBlankValidation('form13','document.form13." & fname & "',true,'\'" & fname & k17 & "\'');"
				' Sadece sayıysa
				If ftip = 3 Then Response.Write "FDK_AddIntegerValidation('form13','document.form13." & fname & "',false,'\'\'','\'\'',true,'\'" & fname & k18 & "\'');"
			End if
		Next
	
	Response.Write "FDK_Validate('form13',false,false,'" & k19 & "\n');return document.MM_returnValue"">"
	Response.Write "<table border=0 cellspacing=1 cellpadding=3 bgcolor=" & td4c & " align=center><tr bgcolor=" & td3c & "><td colspan=2>" & heading13 &"</td></tr>"

'Properties vs.
'**********************
	Set kayittmp = Server.CreateObject("ADODB.Recordset")
	kayittmp.Open td13, conn13, 3, 1, &H0002

	For i = 0 To toplamf
			fname = kayit13(i).Name	
'			fdeger = kayit13(i).Value
			ftip = kayit13(i).Type
			fatt = kayit13(i).Attributes
			fsize = kayit13(i).definedsize
			fpro = "<span style=font-size:10px>"
				If (fatt AND &H00000020) = &H00000020 Then fpro = fpro & "<a href=# title=""" & k20 & """ style=cursor:help> " & penoblankch & " </a>" Else fpro = fpro & "<a href=# title=""" & k21 & """ style=cursor:help> <strike>" & penoblankch & "</strike> </a>"
			fpro = fpro & "</span>"

'**********************
	If pk13 <> fname Then
		Response.Write "<tr bgcolor=" & td5c & "><td align=right><b>" 
		Call dict(fname) 
		Response.Write "</b> : </td>"
			Select Case ftip
					Case 11 'Boolean
							Response.Write "<td><input type=checkbox value=yes name=" & fname & " tabindex=" & i & ">" & fpro & "</td>"
					Case 203 'memo 
						Response.Write "<td><font face=""wingdings"">±</font><a href=""javascript:;"" title=""" & k74 & """ onClick=""ekle('http://','"& fname &"', '" & k42 & " ', '" & k44 & " ')""> <b>http://</b></a> | <font face=""wingdings"">*</font> <a href=""javascript:;"" title=""" & k44 & """ onClick=""ekle('mailto:','"& fname &"', '" & k45 & " ', '" & k43 & " ' )""> <b>" & penoemail & "</b></a> | <font face=""webdings"">Ÿ</font><a href=""javascript:;"" title=""" & k47 & """ onClick=""resimekle('"& fname &"')""> <b>" & k48 & "</b></a> | <a href=""javascript:;"" title=""" & k49 & """ onClick=""tagekle('b','"& fname &"','" & k50 & "' )""><b>B</b></a> | <a href=""javascript:;"" title=""" & k51 & """ onClick=""tagekle('i','"& fname &"','" & k52 & "' )""><b><i>I</i></b></a> | <a href=""javascript:;"" title=""" & k53 & """ onClick=""tagekle('u','"& fname &"','" & k54 & "' )""><b><u>U</u></b></a> | <a href=""javascript:;"" title=""" & k55 & """ onClick=""sonaekle('<hr>','"& fname &"')""><b>Hr</b></a> | <a href=""javascript:;"" title=""" & k56 & """ onClick=""tagekle('quote','"& fname &"','Bloklanacak metni giriniz :' )""> <b>Quote</b> </a><br><textarea name=" & fname & " wrap=virtua tabindex=" & i & "></textarea>" & fpro & "</td></tr>"
					Case 205 ' OLE Nesnesi bmp vs.
						Response.Write "<td>" & k22 & "</td></tr>"	
					Case Else
						Response.Write "<td><input type=text name=" & fname & " tabindex=" & i & ">" & fpro & " </td></tr>"
				End Select
		End if
	Next
	Response.Write "<tr><td colspan=2 align=center bgcolor=" & td3c & "><input type=Submit value=Tamam class=buton tabindex=" & i & "></td></tr></table></form>"
	Response.Write deep
	kayittmp.Close
	Set kayittmp = Nothing
	kayit13.Close
	Set kayit13 = Nothing

End Function

'********************************************
Function EkleOK() 'Yeni Kayıt Ekle İşlem...
'********************************************
	Set kayit13 = Server.CreateObject("ADODB.Recordset")
	kayit13.Open td13, conn13, 1, 2
	toplamf = kayit13.Fields.Count - 1
	
	kayit13.Addnew
	For i = 0 To toplamf
		fname = kayit13(i).Name
		fdeger = kayit13(i).Value
		ftip = kayit13(i).Type
		fatt = kayit13(i).Attributes
		yenideger = Replace(Request.Form(fname), "'","''")		
		If fname <> pk13 Then ' Primary Key ise 
			If  yenideger <> "" Then 'Hiç bir şey girilmediyse
					If ftip <> 205 Then ' Ole nesnesi ise
						Select Case ftip
							Case 11 'Boolean
								If yenideger = "yes" Then yenideger = True Else yenideger = False
							Case 7 ' Tarih ise
								If not IsDate(yenideger) Then yenideger = Null
						End Select
'						Response.Write fname & " - " & yenideger & "<br>"
						kayit13(fname) = yenideger
					End if
				End if
		End If
	Next
	kayit13.Update
	
	kayit13.Close
	Set kayit13 = Nothing
	Response.Redirect busayfa & "?sayfa=" & sayfa 

End Function

'********************************************
Function Sil(kayitno) 'Seçili Kaydı Silme
'********************************************
	conn13.Execute("DELETE * FROM " & td13 & " WHERE " & pk13 & "=" & kayitno)
	conn13.Close
	Set conn13 = Nothing
	Response.Redirect busayfa & "?sayfa=" & sayfa
End Function

'********************************************
Function dbselector() 'Veritabanı seçme
'********************************************
If specialdb <> "" Then 'Select a special db
	If specialodbc <> "" Then 'DSN / ODBC
		Session("db13") = "dsn=" & specialodbc
		If specialpass <> "" Then Session("db13") = specialodbc & ";pwd=" & specialpass
	End if

	If  specialdsnless <> "" Then 'DSNLess
		Session("db13") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(specialdsnless) & ";Persist Security Info=False" 
		If specialpass <> "" Then Session("db13") = specialdsnless & ";pwd=" & specialpass
		Response.Redirect busayfa
	End if
Else
Dim dbtype 'DSN, JET DSNLESS vs...
dbtype = Request.Querystring("dbtype")
If Request.Querystring("db") <> "" Then
Session("dbname") = Request.Querystring("db")
	Select Case dbtype
	Case 1 'DSN
		Session("db13") = "dsn=" & Request.Querystring("db")
		If Request.Querystring("pwd") <> "" Then 	Session("db13") = "dsn=" & Request.Querystring("db") & ";pwd=" & Request.QueryString("pwd")
	Case Else ' Server Mappath / Jet
		Session("db13") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(Request.Querystring("db")) & ";Persist Security Info=False"
		If Request.Querystring("pwd") <> "" Then Session("db13") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(Request.Querystring("db")) & ";Persist Security Info=False"
	End Select
	db13 = Session("db13")
	If Request.Querystring("standart") = "yeah" Then 'Veritabanını standart olarak belirle
		Response.Cookies("db") = db13
		Response.Cookies("db").Expires = Date + 90
	End if
	Response.Redirect busayfa
Else 'Veritabanı adresini gir !...
	Response.Write header13
	Response.Write head13	
	Response.Write "<table width=100% border=0 cellspacing=1 cellpadding=4 	bgcolor=" & td1c & "><form name=""selectdb"" method=GET action=" & busayfa & "><tr><td bgcolor=" & td5c & " align=center width=50% >"
	Response.write "<table><tr><td align=right><b>* Veri tabanı :</b> </td><td><input type=text name=db style=width:140px;font-size:12px value=""test.mdb""></td></tr><tr><td align=right><b>Şifre :</b> </td><td><input type=text name=pwd style=width:140px;font-size:12px></td><tr><td align=right><b>Bağlantı tipi :</b> </td><td><select name=dbtype style=width:140px><option value=1>DSN</option><option value=2 selected>DSNLess/OLEDB</option></select></td><tr><td align=right><input class=clean type=checkbox name=standart value=yeah></td><td><b>Standart olarak belirle</b></td></tr><tr><td>&nbsp;</td><td><input type=submit value=""" & k23 & """ style=width:100px></td></tr></table>"
	Response.Write "</td><td class=P13 width=50% ><font color=#FF0000>peno</font> DB Editor " & penover &"</td></tr></form><tr><td colspan=2 bgcolor=#f7f7f7><b>.01 | DSN</b> " & k24 &" (<b>test.mdb</b> " & penoor & " <b>../db/xxx.mdb</b> - " & k25 & ")<br><b>.01 | DSNless/OLEDB</b> " & k24 & " (<b>test</b> " & penoor & " <b>xxx</b> - gibi)<br><b>.02 | </b>" & k26 & "</br> <b>.03 | </b> " & k27 & "</td></tr></table>"

' Dosyalar arası gezinti / db arama
'********************************************
	Dim dosya, ustklasor, klasorler,folding,foldlen, xjump, xcounter, tdcounter, tdlimit
	xjump = 15
	tdlimit = 3
	tdcounter = 0
	Response.Write "<form name=hepsi method=GET action=" & busayfa & "><table width=100% border=0 cellspacing=1 cellpadding=3 bgcolor=" & td1c & "><tr bgcolor=" & td4c & "><td colspan=" & tdlimit &" bgcolor=" & td3c & "><span class=x1> " & k30 & " </span>(" & k28 & ")</td></tr><tr bgcolor=" & td5c & "><td>"
	folderspec = server.mappath("/" & Request.Querystring("folder"))
	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	If Request.Querystring("folder") <> "" Then  'Üst Klasör Atraksiyonları
			ustklasor = Request.Querystring("folder")
			sayfastr = Split(ustklasor,"/",-1,1)
			ustklasor = ""
			For i = 0 to Ubound(sayfastr)
				If i = Ubound(sayfastr)-1 Then 
					ustklasor = ustklasor + sayfastr(i)
				ElseIf i <> Ubound(sayfastr) Then
					ustklasor = ustklasor + sayfastr(i) & "/"
				End if
			Next
			
		If Ubound(sayfastr) < 1 Then ustklasor = ""
		Response.Write "<a href=" & busayfa & "?folder=" & ustklasor & "><font face=""Webdings"" color=#FF3300>5 </font>" & k29 & "</a><br>"
	End if

	If fso.GetFolder(folderspec) <> "" Then Set fold = fso.GetFolder(folderspec)
			'Klasörlere tıklayabilme / Navigasyon
			klasorler = Split(Request.Querystring("folder"),"/",-1,1) 
			For i = 0 to Ubound(klasorler)
				folding = folding + klasorler(i) + "/"
				foldlen = Len(folding)
				foldlink = Left(folding,foldlen-1)
				Response.Write String(i,Chr(160)) & " <input class=smallput type=checkbox name=""" & foldlink& "fld13666""><a href=" & busayfa & "?folder=" & foldlink & "><font face=""Webdings"" color=#FF3300>4 </font>" & klasorler(i) & "</a><br>"
			Next
			'Navigasyon bitiş
	for each subfolder in fold.subFolders 'listeleme xjump = kaçar kaçar liste	
		If tdcounter = tdlimit AND xcounter = xjump Then
			Response.Write "</td></tr><tr bgcolor=" & td5c & "><td nowrap valign=top>" 
			tdcounter = 0
			xcounter = 0
		Else
			If xcounter = xjump Then 
			Response.Write "</td><td nowrap valign=top>" 
			xcounter = 0
			tdcounter = tdcounter + 1
		End if
		End if
		xcounter = xcounter+1
		If Request.Querystring("folder") <> "" Then		
			Response.Write(" <input class=smallput type=checkbox name=""" & subfolder.name & "fld13666"">&nbsp;&nbsp;&nbsp;<a href=" & busayfa & "?folder=" & Request.Querystring("folder") & "/" & subfolder.name & " title=""" & k31 & """><font face=""Webdings"" color=#FF3300>2 </font>" & subfolder.name & "</a><br>")		
		Else
			Response.Write("<input class=smallput type=checkbox name=""" & subfolder.name & "fld13666""> &nbsp;&nbsp;&nbsp;<a href=" & busayfa & "?folder=" & subfolder.name & " title=""" & k31 & """><font face=""Webdings"" color=#FF3300>2 </font>" & subfolder.name & "</a><br>")
		End if
	Next
	Set objdosyalar = fso.GetFolder(Server.MapPath("/" & Request.Querystring("folder")))
	For Each objDosya in objdosyalar.Files
		dosyaext = fso.GetExtensionName(objDosya.Name)
		If dosyaext = "mdb" Then 
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face=""Webdings"" color=#FF3300> </font><a href=# onClick=""document.selectdb[0].value ='/" & Request.Querystring("folder") & "/" & objDosya.Name & "'"" class=admin>" & objDosya.Name & "</a> | <a href=" & busayfa & "?db=/" & Request.Querystring("folder") & "/" & objDosya.Name & " class=admin> " & k32 & "</a><br>"
		End if
	Next

	set fold = nothing
	set FSO = nothing	

	Response.Write "</td></tr><tr bgcolor=" & td3c & "><td colspan=" & tdlimit & "><input type=radio name=olay value=yedekle class=clean style=width:14px;height:14px checked> <b>" & k33 & " </b><input type=radio name=olay value=sil class=clean style=width:14px;height:14px> <b>" & penodel & " </b> <br><br><input type=submit value=Tamam style=width:100px;background-color:orange> </td></tr></table></form>"
End if
Response.End

End if 'Özel db seçim bitimi
End Function

'********************************************
Function tipacikla (ftip) 'Fieldların açıklamalarını yazdırma
'********************************************
	Select Case ftip
		Case 2
			ftipacikla = "Integer | Sayı"
		Case 3
			ftipacikla = "Long Integer | Uzun sayı "
		Case 4
			ftipacikla = "Single"
		Case 5
			ftipacikla = "Double"
		Case 6
			ftipacikla = "Currency"
		Case 7
			ftipacikla = "Date/Time | Tarih"
		Case 11
			ftipacikla = "Boolean"
		Case 17
			ftipacikla = "Byte"
		Case 7
			ftipacikla = "Date/Time | Tarih"
		Case 203
			ftipacikla = "Memo"
		Case Else
			ftipacikla = "String"
	End Select
	tipacikla = ftipacikla
	
End Function

'********************************************
Function schema13() 'Veritabanının schemasının gösterilmesi tablolar vs.
'********************************************
set kayit13 = conn13.OpenSchema(20)
	do while not kayit13.eof
			if  kayit13("table_type")="TABLE" Then ' Tablo ise
				x = kayit13("table_name")
				Response.Write(x & "<br>")
			End if
		kayit13.movenext
	loop
kayit13.Close
Set kayit13 = Nothing
End Function

'********************************************
Function list13(buyer) 'Veritabanının schemasının gösterilmesi tablolar vs.
'********************************************
set kayitlist13 = conn13.OpenSchema(20)
Response.Write "<form name=""tdlist""><table width=100% border=0 cellspacing=1 cellpadding=4 bgcolor=" & td1c & "><tr><td bgcolor=#f7f7f7 align=center width=250>"
Response.Write "<script language=""JavaScript"">"& vbNewLine
Response.Write "<!--"& vbNewLine
Response.Write "function MM_jumpMenu(targ,selObj,restore){ //v3.0"& vbNewLine
Response.Write "  eval(targ+"".location='""+selObj.options[selObj.selectedIndex].value+""'"");"& vbNewLine
Response.Write "  if (restore) selObj.selectedIndex=0;"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function MM_findObj(n, d) { //v4.01"& vbNewLine
Response.Write "  var p,i,x;  if(!d) d=document; if((p=n.indexOf(""?""))>0&&parent.frames.length) {"& vbNewLine
Response.Write "    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}"& vbNewLine
Response.Write "  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];"& vbNewLine
Response.Write "  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);"& vbNewLine
Response.Write "  if(!x && d.getElementById) x=d.getElementById(n); return x;"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write vbNewLine 
Response.Write "function MM_jumpMenuGo(selName,targ,restore){ //v3.0"& vbNewLine
Response.Write "  var selObj = MM_findObj(selName); if (selObj) MM_jumpMenu(targ,selObj,restore);"& vbNewLine
Response.Write "}"& vbNewLine
Response.Write "function msgx(msg) {"& vbNewLine
Response.Write "  document.MM_returnValue = confirm(msg + ' " & k34 & "');" & vbNewLine
Response.Write "}"& vbNewLine

If Request.Querystring("olay") = "update" OR Request.Querystring("olay") = "insert" Then
	Response.Write "function ekle(deger, ad, qst, qsc) {" & vbNewlline
	Response.Write "var xlink = prompt(qst, '" & k35 & "') ;"
	Response.Write "var xtxt = prompt( qsc , '" & k35 & "') ;"
	Response.Write "var yeni = ""document.form13."" + ad + "".value""; "
	Response.Write "deger = '<a href=""' + deger + xlink + '"">' + xtxt + '</a>'; "
	Response.Write "yeni = eval(yeni) + deger ;"
	Response.Write "var zx = ""document.form13."" + ad + "".value="" + 'yeni' ;"
	Response.Write "eval(zx);"
	Response.Write "}"

	Response.Write "function tagekle(deger, ad, qst) {" & vbNewlline
	Response.Write "var xtag = prompt(qst,'" & k35 & "') ;"
	Response.Write "var yeni = ""document.form13."" + ad + "".value"" ;"
	Response.Write "deger =  '<' + deger + '>' + xtag + '</' + deger + '>' ;"
	Response.Write "yeni = eval(yeni) + deger ;"
	Response.Write "var zx = ""document.form13."" + ad + "".value="" + 'yeni' ;"
	Response.Write "eval(zx);"
	Response.Write "}"

	Response.Write "function sonaekle(deger, ad) {" & vbNewlline
	Response.Write "var yeni = ""document.form13."" + ad + "".value""; "
	Response.Write "yeni = eval(yeni) + deger ;"
	Response.Write "var zx = ""document.form13."" + ad + "".value="" + 'yeni' ;"
	Response.Write "eval(zx);"
	Response.Write "}" & vbnewline

	Response.Write "function resimekle(ad) {" & vbNewlline
	Response.Write "xresim = prompt('" & k36 & "','adres') ;"
	Response.Write "var yeni = ""document.form13."" + ad + "".value""; "
	Response.Write "var deger = '<img src=""' + xresim + '"">' ; "
	Response.Write "yeni = eval(yeni) + deger ;"
	Response.Write "var zx = ""document.form13."" + ad + "".value="" + 'yeni' ;"
	Response.Write "eval(zx);"
	Response.Write "}" & vbnewline
End if 
Response.Write "function focuser() {document.forms[1].elements[1].focus();}"
Response.Write "//-->"& vbNewLine
Response.Write "</script>"

		Response.Write "  <select name=""tdlister"">"

	do while not kayitlist13.eof
			if  kayitlist13("table_type")="TABLE" Then ' Tablo ise
				tdname = kayitlist13("table_name")
				Response.Write "	<option value=" & busayfa &"?olay=tdselect&tablo=" & tdname 
				if  tdname = td13 then
					Response.Write " selected style=background-color:orange"
				End if
				Response.Write ">" 
				Call dict(tdname)
				Response.Write "</option>"
			End if
		kayitlist13.movenext
	loop
			Response.Write "  </select>"
			Response.Write "  <input type=""button"" name=""go"" value=""" & k37 & """ onClick=""MM_jumpMenuGo('tdlister','parent',0)"" style=width:80 px>"
			Response.Write "</td><td colspan=""2""><span class=P13><font color=#FF0000>peno</font> DB Editor " & penover & "</span></td></TR>"
			Response.Write "</table></form>"
	kayitlist13.Close
	Set kayitlist13 = Nothing

End Function

'********************************************
Function tdselector() 'Aktif tablo seçimini değiştirir
'********************************************
Session("aramasql") = ""
if Request.Querystring("tablo") <> "" then
	Session("td13") = Request.Querystring("tablo")
	td13 = Session("td13")
	Session("pk13") = ""
	Session("yon") = ""
	Response.Cookies("peno")("td13") = Session("td13")
	Response.Cookies("peno").Expires = Date + 1000
Else
	If Request.Cookies("peno")("td13") <> "" Then
			Session("td13") = Request.Cookies("peno")("td13")
			td13 = Session("td13")
	Else
		set kayit13 = conn13.OpenSchema(20)
			do while not kayit13.eof
					if  kayit13("table_type")="TABLE" Then ' Sadece tabloları al
						tbname = kayit13("table_name")
						Session("td13") = tbname
						td13 = Session("td13")
						Response.Cookies("peno")("td13") = Session("td13")
						Response.Cookies("peno").Expires = Date + 1000
					End if
				kayit13.movenext
			loop
	End if
End if

' Primary Key kontrolü (autoincrement özelliğine göre !) 
'Otomatik Sayı olan alan Birincil anahtar (Primary Key) olarak kabul edilir
' KEYCOLUMN özelliği sorun çıkatrtabiliyor | Acc 2000
'***************************************
Set kayittmp = Server.CreateObject("ADODB.Recordset")
kayittmp.Open td13, conn13, 3, 1, &H0002

Dim ainc' Autoincrement bulunmazsa otomatik olarak ilk alanı PK yap
ainc = True
for each feld in kayittmp.fields
        if feld.Properties("ISAUTOINCREMENT") = True then
          Session("pk13") = feld.name
		  pk13 = Session("pk13")
		  ainc = False
        end if
 next
	If (ainc) Then
		kayittmp.MoveFirst
		Session("pk13") = kayittmp(0).Name
		pk13 = Session("pk13")
	End if
Call listele()
End Function

'********************************************
Function encoder(deger) ' HTMLEncode  & Yaz
'********************************************
	If deger <> "" Then deger = Server.HTMLEncode(deger)
	Response.Write deger
End Function

'********************************************
Function encoder2(deger) ' HTMLEncode
'********************************************
	If deger <> "" Then deger = Server.HTMLEncode(deger)
End Function

'********************************************
Function P13Trim (p13field,p13lenght,p13trimtype,p13text) ' Uzun kayıtları kısaltır
'********************************************
	if p13field <> "" then
		if p13trimtype = True then
			if LEN(p13field) < p13lenght then
			p13field = RIGHT(p13field, p13lenght)
			Else
			p13field = RIGHT(p13field, p13lenght) & p13text
			End if
		Else
			if LEN(p13field) < p13lenght then
			p13field = LEFT(p13field, p13lenght)
			Else
			p13field = LEFT(p13field, p13lenght) & p13text
			End if
		End if
		Response.Write Server.HTMLEncode(p13field)
	End if
End Function

'********************************************
Function guncelle(kayitno) ' Kayıt Güncelleme
'********************************************

	heading13 = "<span class=x1>" & k39 & "</span>"
	query13 = "SELECT * FROM " & td13 & " WHERE " & pk13 & "=" & kayitno
	Set kayit13 = conn13.Execute(query13)
	'*Field Sayısı
	toplamf = kayit13.Fields.Count - 1
	
	Response.Write header13
	Response.Write head13	
	Call list13( "<a href=" & busayfa & ">" & td13 & "</a> <b>»</b> " & Request.Querystring("no") & k75)
		    If kayit13.EOF And kayit13.BOF Then 
				Response.Write "<table width=100% ><tr><td align=center><b>" & Request.Querystring("no") & "</b> " & k40 &" <b>" 
				Call dict(td13)
				Response.Write "</b> " & k41 & "<br><br><a href=" & Request.ServerVariables("HTTP_REFERRER") & "> " & penoback & " </a></td></tr></table>"
				
				kayit13.Close
				Set kayit13 = Nothing
				Response.End
			End if' Boş ise göster

	Response.Write "<form name=form13 action=" & busayfa & "?sayfa=" & sayfa & "&olay=updateOk&no=" & kayitno & " method=POST>"
	Response.Write "<table border=0 cellspacing=1 cellpadding=3 bgcolor=" & td4c & " align=center><tr bgcolor=" & td3c & "><td bgcolor=#00FF00><a href=" & busayfa & "?olay=delete&no=" & kayitno & " onClick=""msgx('" & kayitno & "');return document.MM_returnValue""><font face=Webdings color=#000000>r</font> " & penodelr & "</a></td><td>" & heading13 & "</td></tr>"
	Dim tabber
	tabber = 0

	For i = 0 To toplamf
		tabber = tabber + 1
		fname = kayit13(i).Name	
		fdeger = kayit13(i).Value
		ftip = kayit13(i).Type
		fatt = kayit13(i).Attributes

		fpro = "<span style=font-size:10px>"
		If (fatt AND &H00000020) = &H00000020 Then fpro = fpro & "<a href=# title=""" & k20 & """ style=cursor:help> " & penoblankch & " </a>" Else fpro = fpro & "<a href=# title=""" & k21 & """ style=cursor:help> <strike>" & penoblankch & "</strike> </a>"
		fpro = fpro & "</span>"

		Response.Write "<tr bgcolor=" & td5c & "><td align=right><b>"
		Call dict(fname)
		Response.Write "</b> : </td>"
			
'Readonly
		If fname = pk13 Then '* Primary Key ise 
			Response.Write "<td><input tabindex=" & tabber & " type=text name=" & fname & " value=""" & fdeger & """></td></tr>"
		Else
			Select Case ftip
				Case 11 'Boolean
						Response.Write "<td><input tabindex=" & tabber & " type=checkbox value=yes name=" & fname
						If fdeger = true then
							Response.Write " checked>"
						Else
							Response.Write "></td>"
						End if
				Case 203 'memo 
					Call encoder2(fdeger)
					Response.Write "<td><font face=""wingdings"">±</font><a href=""javascript:;"" title=""" & k74 & """ onClick=""ekle('http://','"& fname &"', '" & k42 & " ', '" & k44 & " ')""> <b>http://</b></a> | <font face=""wingdings"">*</font> <a href=""javascript:;"" title=""" & k44 & """ onClick=""ekle('mailto:','"& fname &"', '" & k45 & " ', '" & k43 & " ' )""> <b>" & penoemail & "</b></a> | <font face=""webdings"">Ÿ</font><a href=""javascript:;"" title=""" & k47 & """ onClick=""resimekle('"& fname &"')""> <b>" & k48 & "</b></a> | <a href=""javascript:;"" title=""" & k49 & """ onClick=""tagekle('b','"& fname &"','" & k50 & "' )""><b>B</b></a> | <a href=""javascript:;"" title=""" & k51 & """ onClick=""tagekle('i','"& fname &"','" & k52 & "' )""><b><i>I</i></b></a> | <a href=""javascript:;"" title=""" & k53 & """ onClick=""tagekle('u','"& fname &"','" & k54 & "' )""><b><u>U</u></b></a> | <a href=""javascript:;"" title=""" & k55 & """ onClick=""sonaekle('<hr>','"& fname &"')""><b>Hr</b></a> | <a href=""javascript:;"" title=""" & k56 & """ onClick=""tagekle('quote','"& fname &"','Bloklanacak metni giriniz :' )""> <b>Quote</b> </a><br><textarea tabindex=" & tabber & " name=" & fname & " wrap=virtual>" & fdeger & "</textarea>" & fpro & "</td></tr>"
				Case 205 ' OLE Nesnesi bmp vs.
					Response.Write "<td>" & k22 & "</td></tr>"	
				Case Else
					Call encoder2(fdeger)
					Response.Write "<td><input tabindex=" & tabber & " type=text name=" & fname & " value=""" & fdeger & """> " & fpro & "</td></tr>"
			End Select
		End If
	Next
	tabber = tabber + 1
	Response.Write "<tr><td colspan=2 align=center bgcolor=" & td3c & "><input tabindex=" & tabber & " type=SUBMIT value=" & penook & " class=buton></td></tr></table></form>"
	Response.Write deep
	
	kayit13.Close
	Set kayit13 = Nothing

End Function

'********************************************
Function guncelleOk(kayitno)
'********************************************
	Dim uppos ' Update uygunmu ? / update edilemeyecekler için...
	query13 = "SELECT * FROM " & td13 & " WHERE " & pk13 & "=" & kayitno
	Set kayit13 = Server.CreateObject("ADODB.Recordset")
	kayit13.Open query13, conn13, 1, 2
	toplamf = kayit13.Fields.Count - 1
	For i = 0 To toplamf
		fname = kayit13(i).Name
		ftip = kayit13(i).Type
		yenideger = Request.Form(fname)
		uppos = True
		If fname <> pk13 Then 										'PKyi geç
			Select Case ftip
				Case 2, 3, 4, 5, 6, 17									'numberlar
					If not IsNumeric(yenideger) Then yenideger = 0
				Case 11 												'yes/no
					If yenideger = "yes" Then
						yenideger = True
					Else
						yenideger = False
					End If
				Case 7													'tarih
					If not IsDate(yenideger) Then yenideger = Null
				Case 207												' OLE Nesnesi
					uppos = False
				Case Else												'string
					If Trim(yenideger) & "" = "" Then yenideger = Null
			End Select
				If (uppos) Then
					kayit13(fname) = yenideger
				End if
		End If
	Next
	kayit13.Update
	
	kayit13.Close
	Set kayit13 = Nothing

	Response.Redirect busayfa & "?sayfa=" & sayfa

End Function

'********************************************
Function listele() 'Esas olay kayıtları listeleme
'********************************************
	
	Dim sayfasay ' Kaçar kaçar listeleneceği
		if Request.Querystring("sayfasay") <> "" Then 
			Session("sayfasay") = Request.Querystring("sayfasay")
		Else
			Session("sayfasay") = Request.Cookies("penooth")("sayfa")
		End if

		If Session("sayfasay") = "" Then 
			sayfasay = 10
			Session("sayfasay") = sayfasay
		Else
			sayfasay = Session("sayfasay")
		End if
		Response.Cookies("penooth")("sayfa") = Session("sayfasay")
		Response.Cookies("penooth").Expires = Date + 1000
	
	Dim ordtable, ordyon ' Hangi tabloya göre listeleme yapılacağını belirleyen tablo
		ordtable = Request.Querystring("order")
		ordyon = Request.Querystring("yon") ' ASC veya DESC

		if ordtable <> "" AND ordyon <> "" Then ' Liste Seçimini kaydet
			Session("pk13") = ordtable
			Session("yon") = ordyon
		End if
		If Session("pk13") = "" AND Session("yon") = "" Then ' Önceden liste seçimi yoksa
			ordtable = pk13
			Session("pk13") = ordtable
			ordyon = "ASC"
			Session("yon") = ordyon
		Else
			ordtable = Session("pk13")
			ordyon = Session("yon")
		End if	
		
	If Session("aramasql") <> "" Then
		query13 = Session("aramasql") & " ORDER BY " & ordtable & " " & ordyon
	Else
		query13 = "SELECT * FROM " & td13 & " ORDER BY " & ordtable & " " & ordyon
	End if

'================ARAMA SQL
	If Request.Form <> "" Then	
		Dim aramatdsquery, x, aratip,tdstoplam ' Arama SQLini oluşturma
		x = 0
		aramatdsquery = "SELECT * FROM " & td13
		Set aramatds = Server.CreateObject("ADODB.Recordset")
		aramatds.CursorLocation = 3
		aramatds.Open aramatdsquery, conn13, 0, 4
		tdstoplam = aramatds.Fields.Count - 1
		query13 = "SELECT * FROM " & td13
		If Request.Form("aratip") <> "" Then aratip = Request.Form("aratip") Else aratip = "OR"
		For i = 0 to tdstoplam
			If Request.Form(aramatds(i).name) <> "" Then
				If x <> 0 Then
					query13 = query13 & " " & aratip & " " & aramatds(i).name & " LIKE '%" & Replace(Request.Form(aramatds(i).name), "'", "''") & "%' "
				Else
					query13 = query13 & " WHERE " & aramatds(i).name & " LIKE '%" & Replace(Request.Form(aramatds(i).name), "'", "''") & "%' "
				End if
				x = x+1
			End if
		Next
		Session("aramasql") = query13
		aramatds.Close
		Set aramatds = Nothing
	End if

	'****************************
	Set kayit13 = Server.CreateObject("ADODB.Recordset")
	kayit13.CursorLocation = 3
	kayit13.Open query13, conn13, 0, 4

hata(1,0) = err.description
hata(1,1) = err.number

	kayit13.PageSize = sayfasay
	ftoplam = kayit13.Fields.Count - 1
		If clng(sayfa) > kayit13.PageCount Then 'Kayıt yoksa
			sayfa = kayit13.PageCount
			Dim sbos
			sbos = True
		End if
	Response.Write header13
	Response.Write head13
	Call list13(td13 & k76)
	lngMaxsayfas = kayit13.PageCount
if (sbos) then ' Kayıt yoksa
	if x > 0 Then
		Session("aramasql") = ""
		Response.Write "<table width=100% ><tr><td align=center>" & k57 & "<br><a href=" & busayfa & ">" & k58 & "</a><br></td></tr></table>"  
	Else
		Response.Write "<table width=100% ><tr><td align=center>" & k59 & " <br><a class=admin href=" & busayfa & "?olay=insert>" & penonewr & "</a><br></td></tr></table>"  
	End if
Else

	' Navigasyon ve tablo kısa bilgileri
	'************************
	Response.Write "<table width=100% cellpadding=4><tr bgcolor=" & td3c & "><td colspan=3><span class=x1>" & td13 & " </span>( " & k60 & " <b>" & kayit13.RecordCount & " </b> " & k61 & ") - "
	Response.Write "<a class=admin href=" & busayfa & "?sayfa=" 
	Response.Write sayfa
	Response.Write "&olay=insert>" & penonewr & "</a></td></tr><tr><td width=35% nowrap>" & sayfa & " / " & lngMaxsayfas & " &nbsp; "
	If sayfa <> 1 Then	Response.Write "<a href=" & busayfa & "?sayfa=1>" & penofirst & "  </a> &#149;" Else Response.Write "<span class=deactive>" & penofirst & "</span> &#149; "
	If sayfa = 1 Then lngPrevNext = 1 Else lngPrevNext = sayfa - 1
	If  sayfa <> 1 Then Response.Write "<a href=" & busayfa & "?sayfa=" & lngPrevNext & "> Geri </a> &#149; " Else	Response.Write "<span class=deactive> " & penoback & " </span>  &#149; "
	If sayfa = lngMaxsayfas Then lngPrevNext = lngMaxsayfas Else lngPrevNext = sayfa + 1
	If sayfa <> lngMaxsayfas Then Response.Write "<a href=" & busayfa & "?sayfa=" & lngPrevNext & ">" & penonext & " </a> &#149; " Else Response.Write "<span class=deactive> " & penonext & " </span> &#149; "
	If sayfa <> lngMaxsayfas Then Response.Write "<a href=" & busayfa & "?sayfa=" & lngMaxsayfas & ">" & penolast & " </a> &#149; " Else Response.Write "<span class=deactive> " & penolast & " </span>  &#149; " 
	Response.Write "<a href=" & busayfa & "> " & k62 & "</a>"
	If Session("aramasql") <> "" Then Response.Write " | <a href=" & busayfa & "?olay=restart> " & k63 & "</a>"
	Response.Write "</td><form name=sfform action=" & busayfa &" method=GET><td>"

	' Listelenen Kayıt sayısını değiştirme
	'********************************************
	Dim numarator (10)
	Response.Write "  <select name=""sayfasay"" class=kayitlar>"
	For i = 1 to Ubound(numarator)
		numarator(i) = i * 10
					Response.Write "	<option value=" & numarator(i)
				if  Int(Session("sayfasay")) = numarator(i) then
					Response.Write " selected style=""background-color:orange"""
				End if
				Response.Write  " >" & numarator(i) & k73 & " </option>"
	Next
		Response.Write "  </select><input type=""submit"" name=""go"" value=""" & k64 & """ style=width:50 px>"
		Response.Write "</td></form><form name=editpk action=" & busayfa &" method=GET><td align=right><input type=hidden value=update name=olay> " & k65 & " <input type=text name=no style=width:25px;font-size:12px> <input type=submit value=""" & penoeditr & """ style=width:90px></td></form></tr></table>"

'********************************************
	Response.Write "<table cellspacing=1 cellpadding=4 width=100% bgcolor=" & td4c & ">"
	Response.Write "<tr>"
	Response.Write "<td class=x1 bgcolor=orange width=100 align=center>" 
	Call dict(td13)
	Response.Write "</td>"
	For i = 0 To ftoplam
		If kayit13(i).Type > 7 then
			Response.Write "<td class=x1 align=center bgcolor=" & td3c & ">"
		Else
			Response.Write "<td nowrap "
			if kayit13(i).Name = Session("pk13") then Response.Write "width=45 "
			Response.Write "align=right class=x1 bgcolor=" & td3c & ">"
		End If

			' ASC, DESC ORDER Kontrolü ve işaretleri
			'******************************
			if ordtable = pk13 then
				if  kayit13(i).Name = pk13 AND ordyon = "DESC" then
					Response.Write "<a href=" & busayfa & "?order=" & kayit13(i).Name & "&yon=ASC title=""" & k66 & """>" 
					Call dict(kayit13(i).Name)
					Response.Write "<font face=Webdings color=#FF6600 style=font-size:12px>6</font>"	
					ortable = kayit13(i).Name
				Elseif kayit13(i).Name = pk13 Then
					Response.Write "<a href=" & busayfa & "?order=" & kayit13(i).Name & "&yon=DESC title=""" & k67 & """>"
					Call dict(kayit13(i).Name)
					Response.Write "<font face=Webdings color=#FF6600 style=font-size:12px>5</font>"
					ortable = kayit13(i).Name
				End if
			Else
				if  kayit13(i).Name = ordtable AND ordyon = "DESC" Then
					Response.Write "<a href=" & busayfa & "?order=" & kayit13(i).Name & "&yon=ASC title=""" & k66 & """>"
					Call dict(kayit13(i).Name)
					Response.Write "<font face=Webdings color=#FF6600 style=font-size:12px>6</font>"
					ortable = kayit13(i).Name
						Elseif kayit13(i).Name = ordtable Then
						ortable = kayit13(i).Name
					Response.Write "<a href=" & busayfa & "?order=" & kayit13(i).Name & "&yon=DESC title=""" & k67 & """>"
					Call dict(kayit13(i).Name)
					Response.Write "<font face=Webdings color=#FF6600 style=font-size:12px>5</font>"
				End if
			End if
		if ortable <> kayit13(i).Name then
			Response.Write  "<a href=" & busayfa & "?order=" & kayit13(i).Name & "&yon=ASC title=""" & k66 & """>" 
			Call dict(kayit13(i).Name)
		End if
	Response.Write "</a></td>"
	Next
	
	Response.Write "</tr>" & vbCrLf
	
	If kayit13.EOF Then
		Response.Write "<tr><td><a href=" & busayfa & "?sayfa=" & kayit13.PageCount & "&olay=insert>" & penonewr &"</a></td></tr>"
	Else
		'Kayıtları Listele
		intCounter = 0
		kayit13.AbsolutePage = sayfa
		
		For intPager = 1 to sayfasay
			intCounter = intCounter + 1
			Response.Write "<tr "
		If  (IE) then 'Internet Explorer ise
				Response.Write " onMouseOver=""this.bgColor = '#FFF0FF'"" "
		End if

				'Zebra Deseni
				'*******************
				If intCounter Mod 2 <> 0 Then
					Response.Write " onMouseOut=""this.bgColor = '" & td5c & "'"" bgcolor=" & td5c
					Else
					Response.Write " onMouseOut=""this.bgColor = '" & td3c & "'"" bgcolor=" & td3c
				End If
			Response.Write ">"
			Response.Write "<td width=100 align=center nowrap><a href=" & busayfa & "?sayfa=" & sayfa & "&olay=update&no=" & kayit13(pk13) & " title=""" & penoeditr & """ class=admin>" & k68 & "</a> "
			Response.Write " | <a href=" & busayfa & "?sayfa=" & sayfa & "&olay=delete&no=" & kayit13(pk13) & " title=""" & penodelr & """ class=admin onClick=""msgx('" & kayit13(pk13).value & "');return document.MM_returnValue"">" & k69 &"</a></td>" & vbCrLf
			For i = 0 To ftoplam
				varFieldValue = kayit13(i)
				intType = kayit13(i).Type
				if Int(kayit13(i).type) <> 205 then 'Photo / OLE vs. değilse
					if Trim(varFieldValue & "") = "" Then
						Response.Write "<td>&nbsp;"
					Else
						Select Case intType
							Case 2, 3, 4, 5, 17
								Response.Write "<td align=right>" 
								Call encoder(varFieldValue)
							Case 6
								Response.Write "<td align=right>" & FormatCurrency(varFieldValue)
							Case 7
								Response.Write "<td align=right>" & FormatDateTime(varFieldValue, myDates)
							Case Else
								Response.Write "<td>"
								Call P13Trim(varFieldValue,90,false,"...")
						End Select
					End If
				Else
					Response.Write "<td>"
				End if
			Response.Write "</td>"
			Next
			Response.Write "</tr>"
			kayit13.Movenext
			If kayit13.EOF Then Exit For
		Next
	End If

' Arama Sistemi
'*****************
	Dim alantip, alanad
	kayit13.MoveFirst
	Response.Write "<form name=arama action=" & busayfa & " method=POST><tr bgcolor=" & td5c & "><td class=x1 width=100 bgcolor=orange><input type=submit value=""" & k70 & """ style=width:100px></td>"
		For i = 0 To ftoplam
			alantip = kayit13(i).Type
			alanad = kayit13(i).Name
			If alanad <> pk13 Then
				Response.Write "<td align=center> <input type=text style=width:80px name=" & alanad & " value=" & Request.Form(alanad) & ">"
			Else
				Response.Write "<td align=center><select name=aratip style=width:80px><option value=OR> " & penoor & " </option><option value=AND>" & penoand & "</option></select> &nbsp;"
			End if
			Response.Write "</td>"
		Next
	Response.Write "</tr></form>"
'************************

	Response.Write "</table>"
	Response.Write deep
End if

	kayit13.Close
	Set kayit13 = Nothing

End Function

' Hatalar // Sessionlardan dolayı back tuşu çalışmaz
'**********************
If err Then
	Response.Write header13
	Response.Write head13
	Response.Write "<table width=100% border=0 cellspacing=1 cellpadding=4 	bgcolor=" & td1c & "><tr><td bgcolor=" & td5c & " width=50% align=center>"
	For i=0 to Ubound(hata)
		If hata(i,0) <> "" Then
		Response.Write hata(i,0) & "<br> " & k71 & hata(i,1)
		End if
	Next
	Response.Write "<br><a href=" & busayfa & "?x=db>" & k72 & "</a><br>"
	Response.Write "</td><td class=P13 width=50% ><font color=#FF0000>peno</font> DB Editor " & penover & "</td></tr></table>"
	Response.Write deep
End if

If NOT IsNull(conn13) Then
	conn13.Close
	Set conn13 = Nothing
End if
End if
%>