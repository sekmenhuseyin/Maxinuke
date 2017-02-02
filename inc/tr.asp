<%
'MENÜLER
	top_menu1 = "Anasayfa"
	top_menu2 = "Haberler"
	top_menu3 = "Dosyalar"
	top_menu4 = "Makale/Yazýlar"
	top_menu5 = "Üyeler"
	top_menu6 = "Anketler"
	top_menu7 = "Forum"

	top_menu_link1 = "defaultt.asp"
	top_menu_link2 = "News.asp"
	top_menu_link3 = "Programs.asp?action=cats"
	top_menu_link4 = "Articles.asp?action=cats"
	top_menu_link5 = "Members.asp?action=members"
	top_menu_link6 = "Polls.asp?action=polls"
	top_menu_link7 = "Forum.asp"

	menu1 = "Menü"
	menu2 = "Üyeler"
	menu3 = "Anket"
	menu4 = "Üyelik"
	menu5 = "Takvim"
	menu6 = "Dil Seçimi"
	menu7 = "Ýstatistikler"
	arama = "Arama"

'ÜYELER

	m_all = "Hepsi"

'HABERLER (news.asp)

	h_sendfriend = "Bu Haber Arkadaþýna Gönder"
	h_print = "Bu Haberi Yazdýr"
	h_totalvote = "Toplam Oy"
	h_averagepoint = "Ortalama Puan"
	h_from = "Bu Haberin Geldiði Yer"
	h_address = "Bu Haber için Adres"
	h_newsvote ="Puan Verin"
	h_nvmsg = "Bu habere puan verin..."
	h_nextpage = "Sonraki Sayfa"
	h_readed = "okunma"

     'ARKADAÞA GÖNDER
	h_yn = "Ýsminiz"
	h_ym = "E-Mail Adresiniz"
	h_fn = "Arkadaþýnýzýn Ýsmi"
	h_fm = "Arkadaþýnýzýn E-Mail Adresi"

	h_subject = "Haber Tavsiyesi"
	h_msg1 = "isimli arkadaþýnýz bu haberi size göndermek istedi"
	h_msg2 = "Adres"
	h_msg3 = "Daha fazla haber için : "
	h_success = "Haber Arkadaþýnýza Gönderildi..."

'MESAJ KUTUSU

	msg_box = "Gelen Kutusu"
	msg_saved = "Giden Kutusu"
	msg_admin = "Yönetimden Gelenler"
	msg_friends = "Arkadaþlarým"
	msg_new = "Yeni Mesaj Yaz"
	msg_delete_want = "Bu Mesajý Sil"
	msg_welcome = "Hoþgeldin"
	msg_unreaded_msgs = "tane okunmamýþ mesajýnýz var."
	msg_no_unreaed_msgs = "Okunmamýþ Mesajýnýz Yok."
	msg_es = "Boþ Alan"
	msg_us = "Kullandýðýnýz Alan"

'ARKADAÞLARIM

	no_friend = "Listenizde Kimse Yok"
	f_name = "Adý Soyadý"
	f_email = "E-Mail Adresi"
	f_info = "Üye Profili"
	f_msg = "Mesaj Gönder"
	f_delf = "Listeden Sil"
	f_success_del = "Seçtiðiniz Üye Listenizden Baþarýyla Silindi..."
	ps_member = "Lütfen Üye Seçiniz"
	added_once = "Bu Üyeyi Önceden Listenize Eklemiþsiniz"

'MESAJ DETAYLARI
	msg_sender = "Gönderen"
	msg_message = "Mesaj"
	reply_button = "Cevapla"
	msg_subject = "Konu"
	msg_date = "Tarih"

'YENÝ MESAJ
	msg_to = "Gidecek Üye"
	msg_sub = "Baþlýk"
	msg_message = "Mesaj"
	msg_save = "Giden Kutusu'na Kaydet"
	msg_ok = "Gönder"
	msg_del = "Sil"
	msg_empty = "Boþ"
	success_sent = "Mesaj Gönderildi..."

'ÜYE SEVÝYELERÝ

	level1 = "Yönetici"
	level2 = "Dosya Editörü"
	level3 = "Makale/Yazý Editörü"
	level4 = "Haber Editörü"
	level5 = "Forum Yöneticisi"
	level6 = "Forum Moderatörü"
	normal = "Üye"

'TAKVÝM SAYFASI
	clndr_birthday = "Bugün Doðan Üyelerimiz"
	clndr_members = "Bugün Üye Olanlar"
	clndr_files = "Bugün Eklenen Dosyalar"
	clndr_articles = "Bugün Eklenen Makale/Yazýlar"

'HESABINIZ

	ya_havenewmsg = "Merhaba <b>"&Session("Member")&"</b>;<br><b>"&Session("New_Msg_Count")&"</b> adet yeni mesajýn var..."
	ya_topic = "Hesabýnýz"
	ya_selecttheme = "Bir Tema Seçiniz"

	ya_last10news = "Son 10 Haberiniz"

	ya_info = "Bilgilerimi Deðiþtir"
	ya_messages = "Mesajlarým"
	ya_theme = "Tema Seç"
	ya_logout = "Logout/Çýkýþ"

'ÝÇ SAYFALAR

	OU_UserName =  "Kullanýcý Adý"
	OU_LastSession = "Son Oturum Tarihi"
	OU_Browser = "Tarayýcý"
	OU_Workstation = "Ýþletim Sistemi"
	poll_totalvotes = "Toplam Oy"
	no_online = "Hiç Bir Üye Sitede Deðil !"
	member_online = "Online"
	member_offline = "Offline"
	call_message = "Mesaj"
	banned_ip_text = "Siteye giriþiniz Yönetici tarafýndan yasaklanmýþtýr..."
	previous_page = "Önceki Sayfa"
	next_page = "Sonraki Sayfa"
	p_back = "« Geri"
	member_details = "Üye Detaylarý"
	add_fl_success = "Seçtiðiniz Üye Arkadaþ Listenize Baþarýyla Eklendi..."
	poll_details = "Anket Detaylarý"
	lost_memname = "Kullanýcý Adýnýz"
	lost_continue = "Devam »"
	comment = "Yorumunuz"
	comment_new = "Yeni Yorum Eklexxxxxxxxxxxxxxxxxxxxxxxxxx"
	look_details = "Detaylarýna bakmak istediðiniz üyenin Kullanýcý Adý'na týklayýnýz"
	fp_msg = "Þifreniz Mail Adresinize Gönderildi..."
	uye_new_password = "Yeni Þifreniz"

'ÜYE DETAYLARI SAYFASI

	detail_memname = "Kullanýcý Adý"
	detail_name = "Ýsim"
	detail_city = "Þehir"
	detail_job = "Meslek"
	detail_sex = "Cinsiyet"
	detail_age = "Yaþ"
	detail_url = "Web Adresi"
	detail_signature = "Ýmza"
	detail_avatar = "Avatar"
	detail_regdate = "Üyelik Tarihi"
	detail_level = "Ünvan"
	detail_logcount = "Login Sayýsý"
	detail_msgcount = "Mesaj Sayýsý"
	detail_lastlogin = "Son Oturum Tarihi"
	detail_invisiblemail = "Gizli"
	detail_empty = "Yazýlmamýþ"
	detail_status = "Kullanýcý Durumu"
	msg_send = "Mesaj Gönder"

	detail_avatar = "Avatar"
	detail_contact = "Haberleþme"
	detail_about = "Hakkýnda"

'ANA SAYFA ÝLE ÝLGÝLÝ

	top_haber = "Son 5 Haber"
	top_forum = "Son 5 Forum"
	replies_count = "cevap"
	top_sayfa = "Sayfa(lar)"
	top_haberler = "Toplam Haber"

	notice = "DUYURU : "

	uye_kuladi = "Nick"
	uye_password = "Þifre"
	uye_submit = "Giriþ »"
	remember_me = "Beni Hatýrla"
	uye_new = "Yeni Kayýt !"
	uye_lostpwd = "Þifremi Unuttum !"
	uye_kayit = "Kaydet"
	uye_guncelle = "Deðiþiklikleri Kaydet"

	bottom_ziyaretciler = "Ziyaretçi"
	bottom_members = "Üye"
	bottom_total = "Toplam"
	online_mems = "Online Üyeler"
	site_hits = "Site Hitleri"

	news_writer = "Ekleyen"
	news_comments = "Yorumlar"
	news_read = "Devamý »"
	comment_writer = "Ekleyen"
	calender_go = "Git"
	sitedekiler = "Sitedekiler"

'BLOKLAR

	language = "Dil Seçimi"
	lang_button = "Tamam"
	poll = "Anket"
	poll_button = "Oy Gönder"
	poll_katilan = "Katýlýmcý"
	last_member = "Son Üye"
	total_member = "Toplam Üye"
	control_panel = "Yönetim Paneli"
	poll_results = "Sonuçlar"
	poll_others = "Tüm Anketler"

'SAYAÇ
	c_yesterday = "Dün"
	c_today = "Bugün"
	c_total = "Toplam"

'YENÝ PROGRAM EKLEME FORMU
	prg_name = "Programýn Adý"
	prg_info = "Açýklama"
	prg_size = "Dosya Boyutu"
	prg_url = "Download Adresi"
	prg_author = "Telif"
	prg_cat = "Kategori"
	prg_succes = "Göndermiþ olduðunuz program elimize ulaþtý.En kýsa zamanda yetkililer tarafýndan incelenip sitedeki yerini alacaktýr..."

'YENÝ ÜYE KAYDI ÝLE ÝLGÝLÝ

	register_mail = "Mail Adresi"
	register_name = "Ýsim"
	register_question = "Gizli Soru"
	register_answer = "Cevap"
	register_city = "Þehir"
	register_job = "Meslek"
	register_sex = "Cinsiyet"
	register_age = "Yaþ"
	register_url = "Web Adresi"
	register_signature = "Ýmza"
	register_avatar = "Avatar"
	register_mailgoster = "Diðer üyeler mail adresimi görebilir"
	male = "Bay"
	female = "Bayan"
	security_code = "Güvenlik Kodunuz"
	security_code_type = "Güvenlik Kodunuzu Yazýnýz"
	update_title =  "Üye Profili"
	cong = "<b>Tebrikler !!!</b><br><br>Kaydýnýz Baþarýyla Yapýldý...Sað taraftaki giriþ formunu kullanarak giriþ yapabilirsiniz."
	congggg = "<b>Tebrikler !!!</b><br><br>Kaydýnýz Baþarýyla Yapýldý...Lütfen Aktivitasyonunuzu tamamlamak için mail kutunuzu kontrol ediniz."
	cong2 = "<b>Tebrikler !!!</b><br><br>Kayýt bilgileriniz baþarýyla güncellendi..."

	gc_information = "Deðiþtirmek istemiyorsanýz boþ býrakýn"

	pass_change = "Þifre Deðiþtir"
	old_pass = "Eski Þifreniz"
	new_pass = "Yeni Þifreniz"
	confirm_pass = "Þifre Onay"
	succes_saved = "Þifreniz Baþarýyla Deðiþtirildi !"

'HATALAR

	error1 = "Lütfen Kullanýcý Adýnýzý Yazýnýz"
	error2 = "Lütfen Þifrenizi Yazýnýz"
	error3 = "Lütfen Mail Adresinizi Yazýnýz"
	error4 = "Lütfen Ýsminizi Yazýnýz"
	error5 = "Lütfen Gizli Sorunuzu Yazýnýz"
	error6 = "Lütfen Cevabýnýzý Yazýnýz"
	error7 = "Kullanýcý Adý Yanlýþ"
	error8 = "Þifre Yanlýþ"
	error9 = "Bu Kullanýcý Adý daha önce baþkasý tarafýndan seçilmiþ.<BR><BR>Lütfen baþka bir kullanýcý adý belirleyiniz..."
	error10 = "Bu Mail baþka birisi tarafýndan kullanýlmaktadýr.<BR><BR><A HREF=uye_islem.asp?action=lostpass&step=1>Þifremi Unuttum !</A>"
	error11 = "Cevabýnýz Yanlýþ..."
	error12 = "Girdiðiniz Kullanýcý Adý sistemde bulunamadý."
	error13 = "Lütfen Mail Adresinizi Düzgün bir biçimde yazýnýz..."
	error14 = "Anket Oluþturulmamýþ"
	error15 = "Kullanýcý adý <b>15</b> karakterden fazla olmamalýdýr !"
	error16 = "Kullanýcý Adýnda Türkçe Karakterler Kullanmayýnýz !"
	error17 = "Bu ID'ye Ait Bir Üye Yok yada Bulunamadý !"

	sc_err1 = "Lütfen Güvenlik Kodunuzu Yazýnýz"
	sc_err2 = "Güvenlik Kodunuz Doðru Deðil"

	s_error0 = "Lütfen Aramak Ýstediðiniz Kelimeyi Giriniz..."
	s_error1 = "Aradýðýnýz kriter Haberler arasýnda bulunamadý..."
	s_error2 = "Aradýðýnýz kriter Dosyalar arasýnda bulunamadý..."
	s_error3 = "Aradýðýnýz kriter Makale/Yazýlar arasýnda bulunamadý..."

	prg_error1 = "Bu Ýsimde Bir Program Zaten Kayýtlý..."
	prg_error2 = "Lütfen Program Ýsmini Yazýnýz..."
	prg_error3 = "Lütfen Program Açýklamasýný Yazýnýz..."
	prg_error4 = "Lütfen Programýn Dosya Adresini Yazýnýz..."
	prg_error5 = "Lütfen Programýnýn Sahibini Yazýnýz..."
	prg_error6 = "Lütfen Programýn dahil olduðu kategoriyi seçiniz..."

	makale_error1 = "Bu Ýsimde Bir Makale/Yazý Zaten Kayýtlý..."
	makale_error2 = "Lütfen Makale/Yazýnýn Baþlýðýný Yazýnýz..."
	makale_error3 = "Lütfen Makale/Yazýyi Yazýnýz..."
	makale_error4 = "Lütfen Kategori Seçiniz..."

	f_error1 = "Lütfen Konu Baþlýðýný Yazýnýz..."
	f_error2 = "Lütfen Mesajýnýzý Yazýnýz..."
	f_error3 = "Hiç Cevap Yazýlmamýþ..."
	f_error4 = "Baþkasýna ait konuyu düzenleyemezsiniz !"
             f_error5 = "Arka Arkaya Kisa Sürede Konu Acamazsiniz !"
             f_error6 = "Arka Arkaya Kisa Sürede Cevap Veremezsiniz !"

	poll_error0 = "Bu Ankete Daha Önceden Oy Kullanmýþsýnýz !"

	no_art = "Bu Kategoride Kayýtlý Makale/Yazý Yok..."
	no_prg = "Bu Kategoride Kayýtlý Program Yok..."
	no_forum = "Forumda Kayýtlý Baþlýk Yok..."
	no_news = "Kayýtlý Haber Yok..."
	no_topics = "Hiç Mesaj Yok"
	empty_comment = "Lütfen Yorum Yazýnýz..."

	msg_error1 = "Lütfen Mesajýnýzý Yazýnýz..."
	msg_error2 = "Lütfen Gönderilecek Kullanýcýyý Seçiniz..."
	msg_error3 = "Böyle bir Üye Bulunamadý !"

	ms_error1 = "Aradýðýn Kelime ile iliþkili bir Üye bulunamadý..."
	ms_error2 = "Arama Yapabilmek için Lütfen Bir Kriter Yazýnýz"

	pass_err1 = "Lütfen Tüm Alanlarý Doldurunuz..."
	pass_err2 = "Lütfen Eski Þifrenizi Doðru Yazýnýz..."
	pass_err3 = "Lütfen Þifre Onayýný Doðru Yazýnýz..."

	module_eof = "Böyle Bir Modül Yok yada Aktif Deðil"

	nv_errmsg = "Bu habere önceden puan vermiþsiniz..."

	advise_err = "Lütfen Tüm Alanlarý Doldurunuz"

'ÜYE GÝRÝÞÝ ÝLE ÝLGÝLÝ

	uyemenu_profilim = "Profilim"
	uyemenu_msgbox = "Mesaj Kutum"
	uyemenu_uyeara = "Üye Arama"
	uyemenu_logout = "Çýkýþ"

'OKUMA SAYFASI
	a_writer = "Yazar"
	a_hit = "Okunma"


'DOSYALAR SAYFASI
	title_prgs_hit = "Popüler Dosyalar"
	title_prgs_new = "Yeni Eklenenler"
	title_prgs_send = "Dosya Ekle"
	title_prgs_stats = "Ýstatistikler"

	pop5prg = title_prgs_hit
	new5prg = title_prgs_new
	p_more = "Daha..."
	sendprg = title_prgs_send

	vote1 = "1 - Çok Kötü"
	vote2 = "2 - Kötü"
	vote3 = "3 - Orta"
	vote4 = "4 - Güzel"
	vote5 = "5 - Çok Güzel"

	vote_program = "Programý Oylayýn"
	votes = "Ortalama"

	pstats_totalcats = "Toplam Kategori"
	pstats_totalprgs = "Toplam Dosya"
	pstats_totaldownload = "Toplam Ýndirilme"
	pstats_totalwa = "Onay Bekleyen"

'DETAYLAR
	p_author = "Telif"
	p_size = "Boyut"
	p_hit = "Hit"
	p_info = "Açýklama"

'SAYFALA BAÞLIKLARI

	news_title = "Haberler"
	m_title = "Yeni Kullanýcý Kaydý"

'ÜYE ARAMA SAYFASI
	ms_criterion = "Kriter"
	ms_select = "Arama Tipi"
	sby_memname = "Kullanýcý Adýna Göre"
	sby_email = "E-Mail Adresine Göre"
	sby_name = "Ad-Soyada Göre"
	sby_city = "Þehire Göre"
	sby_job = "Mesleðe Göre"
	search_button = "  A R A  "

'FORUM

	forum_nomods = "Yok !"
	fs_topic = "Arama"
	m_writer = "Açan"
	m_topics = "Mesajlar"
	m_answers = "Yanýtlar"
	m_lastpost = "Son Gönderen"
	m_read = "Okunma"

	f_whoonline = "Kimler Online"
	f_pop5 = ""
	fr_pop5 = "Cevap"
	fw_pop5 = "Açan"
	f_stats = "Forum Ýstatistikleri"
	f_new_msg  = "Yeni Konu"
	f_reply_msg = "Cevap Yaz"
	f_locked = "Bu Forum Kilitlenmiþtir"
	m_locked = "Bu Konu Kilitlenmiþtir.Cevap Yazamazsýnýz"

	f_lock = "Kilitle"
	f_unlock = "Kilidi Kaldýr"
	f_edit = "Düzenle"
	f_del = "Sil"
	f_entry = "Yasaðý Kaldýr"
	f_noentry = "Yasakla"

	nm_title = "Baþlýk"
	nm_msg = "Mesaj"

	fr_listsmileys = "Smiley Listesi"

	bf_normal = "Açýk Forum - Son Giriþinizden Ýtibaren Yenilik Yok"
	bf_new = "Açýk Forum - Son Giriþinizden Ýtibaren Yenilik Var"
	bf_locked = "Kilitli Forum"
	bf_noentry = "Giriþ Yasak"

	bm_normal = "Açýk Konu - Son Giriþinizden Ýtibaren Yenilik Yok"
	bm_new = "Açýk Konu - Son Giriþinizden Ýtibaren Yenilik Var"
	bm_hot = "Ateþli Konu"
	bm_locked = "Kilitli Konu"

	mem_msg_count = "Mesaj Sayýsý"
	mem_age = "Yaþ"
	mem_sex = "Cinsiyet"
	mem_status = "Kullanýcý Durumu"

	secret_forumcodes = "[B]Kalýn[/B]<br>[I]Ýtalik[/I]<br>[U]Altý Çizgili[/U]<br>[CENTER]Yazý Ortala[/CENTER><br>[URL]"&Session("SiteAddress")&"[/URL]<br>[URL="&Session("SiteAddress")&"]Linkim[/URL]<br>[EMAIL=ad@soyad.com]Mail[/EMAIL]"

	'ARAMA SAYFASI
	fs_word = "Aranacak Kelime"
	fs_forums = "Aranacak Forum"
	fs_button = "Ara !"
	fs_result_1 = "Arama sonucunda yazdýðýnýz sözcükle alakalý "
	fs_result_2 = " tane konu bulundu."



	makale_succes = "Göndermiþ olduðunuz makale/yazý elimize ulaþtý.En kýsa zamanda yetkililer tarafýndan incelenip Makale/Yazýlar arasýndaki yerini alacaktýr..."
%>