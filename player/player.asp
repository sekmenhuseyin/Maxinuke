<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=1254">
<title>PenO</title>
<link rel="stylesheet" rev="stylesheet" type="text/css" href="skin.css">
<link rel="stylesheet" rev="stylesheet" type="text/css" href="slider.css">
</head>
<body scroll="no" oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<object id="wmp" classid="CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6">
</object>
<table width="100%" border="0" cellspacing="0" cellpadding="0" id="table_tab" alwaysview="false">
		<tr>
				<td align="center" background="../images/player/shade_bg.gif"><img src="../images/player/shade_open.gif" width="99" height="8" border="0" class="hand" onclick="toggleTab();"></td>
		</tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0" id="table_player">
		<tr>
				<td width="25" background="../images/player/bg1.gif">
                <p align="center">&nbsp;</td>
				<td width="25" background="../images/player/bg1.gif"><img src="../images/player/btn_open.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_open','','../images/player/btn_open_c.gif',1)" alt="Liste Aç" name="img_open" id="img_open" width="25" height="30" border="0" class="hand" onclick="openPlaylist();"></td>
				<td width="23" background="../images/player/bg1.gif"><img src="../images/player/btn_play.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_play','','../images/player/btn_play_c.gif',1)" alt="Başlat" name="img_play" width="23" height="30" border="0" class="hand" onclick="aPlay();"><img src="../images/player/btn_pause.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_pause','','../images/player/btn_pause_c.gif',1)" alt="Duraklat" name="img_pause" width="23" height="30" border="0" style="display:none;" class="hand" onclick="aPause();"></td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_stop.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_stop','','../images/player/btn_stop_c.gif',1)" alt="Dur" name="img_stop" width="21" height="30" border="0" class="hand" onclick="aStop();"></td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_prev.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_prev','','../images/player/btn_prev_c.gif',1)" alt="Önceki Parça" name="img_prev" width="21" height="30" border="0" class="hand" onclick="aPrev();"></td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_next.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_next','','../images/player/btn_next_c.gif',1)" alt="Sonraki Parça" name="img_next" width="21" height="30" border="0" class="hand" onclick="aNext();"></td>
				<td background="../images/player/bg1.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed;">
								<tr>
										<td width="11"><img src="../images/player/lcd_left.gif" width="11" height="30"></td>
										<td width="50" align="center" nowrap background="../images/player/lcd_bg.gif" class="padd_t2" title="Time Display (Cilck to toggle elapsed/remaining)"><div id="text_duration" name="text_duration" class="shadow bold" onclick="setDurationType();">&nbsp;</div></td>
										<td nowrap background="../images/player/lcd_bg.gif" class="padd_t2"><div id="text_event" name="text_event" class="shadow bold"></div><div id="text_title" name="text_title" class="shadow bold">&nbsp;</div></td>
										<td width="60" align="right" nowrap background="../images/player/lcd_bg.gif" class="padd_t2"><div id="text_bitrate" name="text_bitrate" class="shadow bold">&nbsp;</div></td>
										<td width="11"><img src="../images/player/lcd_right.gif" width="10" height="30"></td>
								</tr>
						</table></td>
				<td width="65" background="../images/player/bg1.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
								<tr>
										<td width="11"><img src="../images/player/lcd_left.gif" width="11" height="30"></td>
										<td align="center" background="../images/player/lcd_bg.gif"><img src="../images/player/btn_shuffle_off.gif" alt="Karışık mod" name="img_shuffle" width="14" height="14" border="0" align="absmiddle" class="hand" onclick="setShuffle();"></td>
										<td align="center" background="../images/player/lcd_bg.gif"><img src="../images/player/btn_loop_on.gif" alt="Yenile" name="img_loop" width="14" height="14" align="absmiddle" class="hand" onclick="setLoop();"></td>
										<td width="11"><img src="../images/player/lcd_right.gif" width="10" height="30"></td>
								</tr>
						</table></td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_rew.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_rew','','../images/player/btn_rew_c.gif',1)" alt="Geri sar" name="img_rew" width="21" height="30" border="0" class="hand" onclick="aREW();"></td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_ff.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_ff','','../images/player/btn_ff_c.gif',1)" alt="İleri sar" name="img_ff" width="21" height="30" border="0" class="hand" onclick="aFF();"></td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_mute_off.gif"onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_mute_off','','../images/player/btn_mute_off_c.gif',1)" alt="Ses kapat" name="img_mute_off" width="21" height="30" border="0" class="hand" onclick="setMute();"><img src="../images/player/btn_mute_on.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_mute_on','','../images/player/btn_mute_on_c.gif',1)" alt="Ses aç" name="img_mute_on" width="21" height="30" border="0" class="hand" onclick="setMute();" style="display:none;"></td>
				<td width="60" align="center" background="../images/player/slider_bg.gif">
						<div class="slider" id="volume" tabIndex="1">
                          <input class="slider-input" id="volumeSlider" size="20"></div>
				</td>
				<td width="21" background="../images/player/bg1.gif"><img src="../images/player/btn_shade.gif" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img_hide','','../images/player/btn_shade_c.gif',1)" alt="Medya bar'ı gizle" name="img_hide" width="21" height="30" border="0" class="hand" onclick="toggleTab();"></td>
		</tr>
</table>
<script language="javascript" type="text/javascript" src="slider.js"></script>
<script language="javascript" type="text/javascript" src="skin.js?<?=time()?>"></script>
</body>
</html>