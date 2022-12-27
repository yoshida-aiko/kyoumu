<!--#include file="../../com_All.asp"-->
<%

m_stoday = gf_YYYY_MM_DD(date(),"/")
	m_iyear = Year(m_stoday)
	m_imonth = Month(m_stoday)-1
m_iMondiff = Request("numDiff")
if m_iMondiff = "" then m_iMondiff = 0
%>

<html>
<head>
<title>カレンダー</title>
</head>
<script language="JavaScript"><!--
/* use free!! , TAB = 4 */
/* ↓一応、名前書いてはありますが、著作権は放棄します。御自由にお使いください。 */
/* ↓使う時は消していいです。(笑) */

// by P,H,M (すい)
// modified by Mugi
// original=http://www.wakusei.ne.jp/~yuuki/web/test/00456.html
//
// 以下、順 早い者勝ち(笑)
//
// thank's いいじまさん		一番はじめの恥ずかしい虫
// thank's 京極どぉ？さん	一番はじめの恥ずかしい虫
// thank's おじさんさん		ネスケの件
// thank's JYA^2さん		閏年処理
// thank's Three-eyeさん	休日処理全般、その他色々(^^)
// thank's NobMochiさん		動作確認
// thank's Jogaさん			「春分の日」「秋分の日」

function mz_calendar(mondiff){
	var		j,z;
	var		today,display;
	var		year,mon,date,day;
	var		nday;
	var		topday;	// 月の最初の曜日
	var		endd;	// 月の最終の日にち
	var		wk;		// 第?週
	var		ln;		// カレンダー表示上の?段
	var		furikae	// 振替休日処理用フラグ。1なら次の日は振替休日
	var		VernalEquinoxDayd  ,VernalEquinoxDay;	// 春分の日
	var		AutumnalEquinoxDayd,AutumnalEquinoxDay;	// 秋分の日

	/***** 表示する文字の属性定義 *****/
	var		COLOR  =   7;		// 0000.0111 MASK 8色
	var		ATTRIB = 248;		// 1111.1000 MASK 6属性
	var		NORM   =   0;		// 000	標準色(色指示無しの文字色)
	var		RED    =   1;		// 001	赤色
	var		BLUE   =   2;		// 010	青色
//	var		       =   3;		// 011
//	var		       =   4;		// 011
//	var		       =   5;		// 011
//	var		       =   6;		// 011
//	var		       =   7;		// 011
	var		UL     =   8;		// 下線
	var		SO     =  16;		// 取消線
	var		IT     =  32;		// イタリック
//	var		       =  64;		// 予備
//	var		       = 128;		// 予備
//	var		       = 256;		// 予備
	now = new Date("<%=m_stoday%>");
	
	
	year = now.getYear();	if( year<1900 ) year+=1900;
	
	mon  = now.getMonth()+1+mondiff;
	while( 12<mon ){	year++;		mon-=12;	}
	while( mon< 1 ){	year-=1;	mon+=12;	}
	date = now.getDate();
	
	nday  = now.getDay();
	if( nday==0 )	day = "日";
	if( nday==1 )	day = "月";
	if( nday==2 )	day = "火";
	if( nday==3 )	day = "水";
	if( nday==4 )	day = "木";
	if( nday==5 )	day = "金";
	if( nday==6 )	day = "土";

	// 今月の最初の曜日を topday に得る。
	topd = new Date(year,mon-1,1);	// 今月の 1日の情報を topd に得る。
//	topd = new Date("<%=m_iyear%>",<%=m_imonth%>,1);	// 今月の 1日の情報を topd に得る。
	topday = topd.getDay();
	// 今月の最終日を endday に得る。 ↓NC4 では動作しませんでした。NN2,3,IE3,4,5なら OK なのにぃ(;_;)
	//	endd = new Date(year,mon,0);	// 今月の最終日の情報を endd に得る。
	//	endday = endd.getDate();		// endday = 今月の最終日付

	if( mon==2 ){
		if( (year%400)==0 )			endday = 29;	/* 400 で割切れる年は2/29まで */
		else if( (year%100)==0 )	endday = 28;	/* 100 で割切れる年は2/28まで */
		else if( (year%  4)==0 )	endday = 29;	/*   4 で割切れる年は2/29まで */
		else						endday = 28;	/* その他の年は 2/28 まで */	
	}																			
	else if( (mon==4)||(mon==6)||(mon==9)||(mon==11) )	endday = 30;				
	else	endday = 31;															

	/***** ↓年によって変化する「春分の日」「秋分の日」を簡易計算により求める *****/
	// 参考：http://www.top.or.jp/~cpop/syunbun.htm
	// 1900～2099年の範囲限定（2100年になったら、責任をもって修正してください。(^^; ）

	// 正式な日付は“前の年の2月1日付けの官報で公示される”ことになっています。
	// 簡易計算によって求めた値が合わない年が発生したら、その年について、適宜、例外処理して下さい。

	VernalEquinoxDayd   = Math.floor(0.24242*year - Math.floor(year/4) + 35.84);	// year年の者分の日
	AutumnalEquinoxDayd = Math.floor(0.24204*year - Math.floor(year/4) + 39.01);	// year年の秋分の日

	// ↓例外処理の例
	//	if( year==20XX ){	// 20XX年なら
	//		VernalEquinoxDayd   = 21;	// 春分の日 = 21日(3月)
	//		AutumnalEquinoxDayd = 23;	// 秋分の日 = 23日(9月)
	//	}
	//	else{	// その他の年なら簡易計算値のままで OK
	//		VernalEquinoxDayd   = Math.floor(0.24242*year - Math.floor(year/4) + 35.84);	// year年の者分の日
	//		AutumnalEquinoxDayd = Math.floor(0.24204*year - Math.floor(year/4) + 39.01);	// year年の秋分の日
	//	}

	VernalEquinoxDay   = 3+VernalEquinoxDayd  /100;		// year年の春分の日
	AutumnalEquinoxDay = 9+AutumnalEquinoxDayd/100;		// year年の秋分の日
	/***** ↑「春分の日」「秋分の日」ここまで↑ *****/

	/* 以下「キリキリ書けやぃ！！」 */
		pre = <%=m_iMondiff%> -1;
		nxt = <%=m_iMondiff%> +1;
		document.write("<pre><tt>");
		document.write("<a href='javascript:nextcal("+pre+");' onclick='nextcal("+pre+")'><<</a>  ");
//		if( mon==3 )		document.write( "   "    + year +"/"+ mon+" 春分" + VernalEquinoxDayd   + "日" );
//		else if( mon==9 )	document.write( "   "    + year +"/"+ mon+" 秋分" + AutumnalEquinoxDayd + "日" );
//		else{
							if( mon<10  )		document.write(" ");	// スペースの調整
							document.write( "   " + year +"/"+ mon +"   " );
//		}
		document.writeln("  <a href='javascript:nextcal("+nxt+");' onclick='nextcal("+nxt+")'>>></a>");
//		document.writeln("<table border='0' width='100%'><tr>");
//		document.writeln("<td align='left'><a href='#' onclick='nextcal("+pre+")'><<</a></td>");
//		document.writeln("<td align='right'><a href='#' onclick='nextcal("+nxt+")'>>></a></td>");
//		document.writeln("</tr></table>");

	document.writeln("");
	document.write('<font color="#FF0000">日<\/font> 月 火 水 木 金 <font color="#0000FF">土<\/font><br>');
	for( j=0 ;j<topday;j++)	document.write("   ");

	z=j+1;	// z  = 曜日(1=日/2=月...7=土)
	ln = 1;

	for( j=1 ; j<=endday ; j++,z++ ){		// ←１ヶ月表示のループ
		wk = 1;					// 第1週
		if(  8<=j )	wk=2;		// 第2週
		if( 15<=j )	wk=3;		// 第3週
		if( 22<=j )	wk=4;		// 第4週
		if( 29<=j )	wk=5;		// 第5週

		// 以降↓ j：日付 / z：曜日(1=日/2=月...7=土) / wk：第?週 / ln：カレンダー表示上の?段目
		//        year：表示する年 / mon：表示する月 / date：今日の日にち

		today = eval(mon+"+"+(j/100));		// today = 月.日

		/***** ここから休日表示処理 *****/
		display = 0;	// 日付の数字の文字色・文字属性をリセットしておく。

		// 振替休日処理(※月末に「国民の休日」は無いという前提で作られています。)
		if( furikae )	display = RED|UL;	// 前日が「国民の祝日」で「日曜日」だったら→お休み。
											// 振替休日＝赤色＆下線
		// ↓国民の祝日
		if( today== 1.01 )	display = RED;	// 元日
		if( today== 2.11 )	display = RED;	// 建国記念の日
		if( today== 4.29 )	display = RED;	// みどりの日
		if( today== 5.03 )	display = RED;	// 憲法記念日
		if( today== 5.05 )	display = RED;	// こどもの日
		if( today== 7.20 )	display = RED;	// 海の日
		if( today== 9.15 )	display = RED;	// 敬老の日
		if( today==11.03 )	display = RED;	// 文化の日
		if( today==11.23 )	display = RED;	// 勤労感謝の日
		if( today==12.23 )	display = RED;	// 天皇誕生日
		if( today == VernalEquinoxDay   )	display = RED;	// 春分の日
		if( today == AutumnalEquinoxDay )	display = RED;	// 秋分の日

		// 「成人の日」「体育の日」処理
		if( 2000>year ){		// 2000年より前
			if( today== 1.15 )	display = RED;	// 成人の日
			if( today==10.10 )	display = RED;	// 体育の日
		}
		else{					// 2000年から
			if( ( mon==1  )&&(z==2)&&(wk==2) ){	display = RED;	}	// 成人の日( 1月第2週の月曜日)
			if( ( mon==10 )&&(z==2)&&(wk==2) ){	display = RED;	}	// 体育の日(10月第2週の月曜日)
		}

		// 振替休日処理
		// ※月末に「国民の休日」は無い（1日が振替休日になることがない）という前提で作られています。
		if( ( display )&&( z<=1 ) )	furikae = 1;	// 今日が「国民の祝日」で「日曜日」だったら次の日は休み
		else						furikae = 0;

		if( today== 5.04 )	display = RED;	// 「憲法記念日」と「こどもの日」に挟まれているから休日
		// “「国民の祝日」に挟まれた1日は休日とする”に該当するのは現在この日だけ。(多分)


		// ユーザーカスタマイズ領域
		if( z<=1 )			display = 1;	// 日曜日
//		if( z==7 )			display = 2;	// 土曜日		土曜日を青で表示したい時
		if( z==7 )			display = 1;	// 土曜日		土曜日を赤で表示したい時

		if( year==1998 ){						// 1998年休日データ
			if( today== 1.02 )	display = RED;		// 休日
			if( today== 1.17 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today== 2.14 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today== 4.30 )	display = RED|UL;	// 振替			赤色＆下線
			if( today== 5.01 )	display = RED;		// ??????
			if( today== 7.27 )	display = RED;		// 夏休み
			if( today== 7.28 )	display = RED;		// 夏休み
			if( today== 7.29 )	display = RED;		// 夏休み
			if( today== 7.30 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 7.31 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 8.14 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 9.19 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today==11.07 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today==11.28 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today==12.29 )	display = RED;		// 休日
			if( today==12.30 )	display = RED;		// 休日
			if( today==12.31 )	display = RED;		// 休日
		}
		if( year==1999 ){						// 1999年休日データ
			if( today== 1.09 )	display = BLUE;		// 土曜日
			if( today== 1.04 )	display = RED;		// 休日
			if( today== 3.22 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 4.30 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 7.20 )	display = NORM|UL;	// 振り替え		標準色＆下線
			if( today== 7.26 )	display = RED;		// 夏休み
			if( today== 7.27 )	display = RED;		// 夏休み
			if( today== 7.28 )	display = RED;		// 夏休み
			if( today== 7.29 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 7.30 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 8.13 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 8.16 )	display = RED|UL;	// 振り替え		赤色＆下線
			if( today== 9.18 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today==10.16 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today==11.06 )	display = BLUE|UL;	// 土曜日		青色＆下線
			if( today==12.29 )	display = RED;		// 休日
			if( today==12.30 )	display = RED;		// 休日
			if( today==12.31 )	display = RED;		// 休日
		}
		if( year==2000 ){						// 2000年休日データ
			if( today== 1.03 )	display = RED;		// 休日
			if( today== 1.04 )	display = RED;		// 休日
			if( today== 1.15 )	display = BLUE|UL;	// 土曜日			青色＆下線
		}


		if( display & UL )				document.write('<u>');							// 下線
		if( display & SO )				document.write('<s>');							// 取消線
		if( display & IT )				document.write('<i>');							// イタリック
		if( (display&COLOR)==RED  )		document.write('<font color="#FF0000">');		// 赤
		if( (display&COLOR)==BLUE )		document.write('<font color="#0000FF">');		// 青
		if( (mondiff==0)&&(j==date) )   document.write("<FONT style='background:#55FF55'>");	// 今日

		if( j<10 )	document.write(" ");		// 日付が1桁ならスペースを1個書く
		document.write('<a href=# class=datelink onclick="inputdate('+year+','+mon+','+j+');return false">'+j+'</\a>');						// 日付を書き書き "φ(- -;)

		if( (mondiff==0)&&(j==date) )	document.write("<\/FONT>");						// 今日
		if( (display&COLOR)==BLUE )		document.write('<\/font>');						// 青
		if( (display&COLOR)==RED  )		document.write('<\/font>');						// 赤
		if( display & IT )				document.write('<\/i>');						// イタリック
		if( display & SO )				document.write('<\/s>');						// 取消線
		if( display & UL )				document.write('<\/u>');						// 下線

		if( 6<z ){	z=0;	ln++;	document.write("<br>");	}	// 週終わり→改行
		else						document.write(" ");		// 日付間のスペース
	}
				document.write("<\/tt><\/pre>");
}

function mz_clock(){
	var hour,mini,sec;

	now = new Date();
	hour = now.getHours();		if( hour<10 )	hour = "0"+hour;
	mini = now.getMinutes();	if( mini<10 )	mini = "0"+mini;
	sec  = now.getSeconds();	if( sec <10 )	sec  = "0"+sec ;

//  form の name↓ ↓input の name
	document.clock.time.value = hour + ":" + mini + "," + sec ;
	setTimeout("mz_clock()",500);	// 次回実行は 500/1000(=0.5)秒後
}

function inputdate(y,m,d){
if(m<10)m="0"+m
if(d<10)d="0"+d
opener.document.frm.<%=Request("txtDay")%>.value=(y+"/"+m+"/"+d);
window.close();
}

function nextcal(pDiff){
wUrl = "default.asp?txtDay=<%=Request("txtDay")%>&numDiff="+pDiff
 document.location.href = wUrl;
 
}
// --></script>

<style><!--
a.datelink{text-decoration:none}
--></style>
<body onLoad="window.focus();">
<center>
<table CellSpacing=0 border=1><tr height=135 align="left" valign="top">
	<td BGcolor="#FFFFFF" height=135>
		<script language="JavaScript"><!--
			mz_calendar(<%=m_iMondiff%>);
		// --></script>
	</td>
</tr></table><font size="1"><br></font>
<input type="button" class="button" value="閉じる" onClick="window.close()">
</center>
</body>
</html>
