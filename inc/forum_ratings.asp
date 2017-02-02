<%
IF f_mem_level = "0"  Then
	IF f_mem_msg < 10 Then
		fm_level = "Er"
		fm_img = "0"
	ElseIF f_mem_msg >= 10 AND f_mem_msg < 20 Then
		fm_level = "Onbaþý"
		fm_img = "1"
	ElseIF f_mem_msg >= 20 AND f_mem_msg < 30 Then
		fm_level = "Çavuþ"
		fm_img = "2"
	ElseIF f_mem_msg >= 30 AND f_mem_msg < 40 Then
		fm_level = "Uzman Onbaþý"
		fm_img = "3"
	ElseIF f_mem_msg >= 40 AND f_mem_msg < 50 Then
		fm_level = "Uzman Çavuþ"
		fm_img = "4"
	ElseIF f_mem_msg >= 50 AND f_mem_msg < 60 Then
		fm_level = "Astsubay Çavuþ"
		fm_img = "5"
	ElseIF f_mem_msg >= 60 AND f_mem_msg < 70 Then
		fm_level = "Astsubay Kýdemlý Çavuþ"
		fm_img = "6"
	ElseIF f_mem_msg >= 70 AND f_mem_msg < 80 Then
		fm_level = "Astsubay Üstçavuþ"
		fm_img = "7"
	ElseIF f_mem_msg >= 80 AND f_mem_msg < 90 Then
		fm_level = "Astsubay Kýdemli Üstçavuþ"
		fm_img = "8"
	ElseIF f_mem_msg >= 90 AND f_mem_msg < 100 Then
		fm_level = "Astsubay Baþçavuþ"
		fm_img = "9"
	ElseIF f_mem_msg >= 100 AND f_mem_msg < 110 Then
		fm_level = "Astsubay Kýdemli Baþçavuþ"
		fm_img = "10"
	ElseIF f_mem_msg >= 110 AND f_mem_msg < 120 Then
		fm_level = "Asteðmen"
		fm_img = "11"
	ElseIF f_mem_msg >= 120 AND f_mem_msg < 130 Then
		fm_level = "Teðmen"
		fm_img = "12"
	ElseIF f_mem_msg >= 130 AND f_mem_msg < 140 Then
		fm_level = "Üsteðmen"
		fm_img = "13"
	ElseIF f_mem_msg >= 140 AND f_mem_msg < 150 Then
		fm_level = "Yüzbaþý"
		fm_img = "14"
	ElseIF f_mem_msg >= 150 AND f_mem_msg < 160 Then
		fm_level = "Binbaþý"
		fm_img = "15"
	ElseIF f_mem_msg >= 160 AND f_mem_msg < 170 Then
		fm_level = "Yarbay"
		fm_img = "16"
	ElseIF f_mem_msg >= 170 AND f_mem_msg < 180 Then
		fm_level = "Albay"
		fm_img = "17"
	ElseIF f_mem_msg >= 180 AND f_mem_msg < 190 Then
		fm_level = "Tuðgeneral"
		fm_img = "18"
	ElseIF f_mem_msg >= 190 AND f_mem_msg < 200 Then
		fm_level = "Tümgeneral"
		fm_img = "19"
	ElseIF f_mem_msg >= 200 AND f_mem_msg < 210 Then
		fm_level = "Korgeneral"
		fm_img = "20"
	ElseIF f_mem_msg >= 210 AND f_mem_msg < 220 Then
		fm_level = "Orgeneral"
		fm_img = "21"
	ElseIF f_mem_msg >= 220 AND f_mem_msg < 230 Then
		fm_level = "Genelkurmay"
		fm_img = "22"
	ElseIF f_mem_msg >= 230 Then
		fm_level = "Mareþal"
		fm_img = "23"
	End IF


ELSE

	IF f_mem_level = "1" Then
		fm_level = level1
		fm_img = "24"
	ElseIF f_mem_level = "2" Then
		fm_level = level2
		fm_img = "11"
	ElseIF f_mem_level = "3" Then
		fm_level = level3
		fm_img = "11"
	ElseIF f_mem_level = "4" Then
		fm_level = level4
		fm_img ="11"
	ElseIF f_mem_level = "5" Then
		fm_level = level5
		fm_img = "11"
	End IF

End IF

fm_include_img = "<img src=""images/yildiz_"&fm_img&".gif"" border=""0"">"
fm_include_img2 = "<img src=""images/rutbe_"&fm_img&".gif"" border=""0"">"
%>