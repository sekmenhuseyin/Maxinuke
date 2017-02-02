var skin_name     = "styx_wmp9_red2";
var skin_by       = "styx";
var skin_email    = "aquamp (at) bystyx.com";
var skin_homepage = "http://bystyx.com";
document.writeln('<script language="javascript" type="text/javascript" src="player.config.js"></script>');
document.writeln('<script language="javascript" type="text/javascript" src="player.js"></script>');
function afmObj_volume(val) {
	if(typeof(vSlider) != "undefined") {
		vSlider.setValue(val);
		showTextEvent("Ses düzeyi - "+val);
	}
}

var showTextEventsT;
function showTextEvent(val) {
	clearTimeout(showTextEventsT);
	if(val) {
		text_title.style.display = "none";
		text_event.style.display = "block";
		text_event.innerHTML = val;
		showTextEventsT = window.setTimeout("showTextEvent()",1200);
	}	else {
		text_title.style.display = "block";
		text_event.style.display = "none";
	}
}

function afmObj_play(a) {
	if(a == 1) {
		img_play.style.display = "none";
		img_pause.style.display = "block";
	} else {
		img_play.style.display = "block";
		img_pause.style.display = "none";
	}
}

function afmObj_shuffle(a) {
	if(a == 1) {
		showTextEvent("Karýþýk mod - Açýk");
		img_shuffle.src = "../images/player/btn_shuffle_on.gif";
	} else {
		showTextEvent("Karýþýk mod - Kapalý");
		img_shuffle.src = "../images/player/btn_shuffle_off.gif";
	}
}

function afmObj_loop(a) {
	if(a == 1) {
		showTextEvent("Yenile - Açýk");
		img_loop.src = "../images/player/btn_loop_on.gif";
	} else {
		showTextEvent("Yenile - Kapalý");
		img_loop.src = "../images/player/btn_loop_off.gif";
	}
}

function afmObj_mute(a) {
	if(a == 1) {
		showTextEvent("Ses - Kapalý");
		img_mute_off.style.display = "none";
		img_mute_on.style.display = "block";
	} else {
		showTextEvent("Ses - Açýk");
		img_mute_off.style.display = "block";
		img_mute_on.style.display = "none";
	}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}