<HTML>
<HEAD>
	<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=Shift_JIS">
<TITLE>Calender</TITLE>
	<input type="hidden" name="hid_txtname" value="<%=Request.QueryString("txtName")%>">

<script language="JavaScript" src="../Common/CommonJS.js"></script>
<script language="javascript">
<!--
// ウィンドウに名前をつける
window.name = "js_calender";

// 呼ばれたプログラムをwURLにいれる
var wDefURL = opener.location.href;
    wDefURL.match(/.asp/);
    wDefURL = RegExp.leftContext;

var cWeek = new Array(7);
var sDate = new Date();

var thisYear = sDate.getYear()
if( thisYear < 2000){	thisYear += 1900 };
var thisMonth = sDate.getMonth()+1;
var thisDate = sDate.getDate();

var toDate= new Date();
var toYear = toDate.getYear()
if( toYear < 2000){	toYear += 1900 };
var toMonth = toDate.getMonth()+1;
var toDay = toDate.getDate();
var todayText;
var m_Sun;
var m_Mon;
var m_Tus;
var m_Wed;
var m_Thu;
var m_Fri;
var m_Sat;
var m_Coment;

// Calender Information
	cWeek[0]="（日）";
	cWeek[1]="（月）";
	cWeek[2]="（火）";
	cWeek[3]="（水）";
	cWeek[4]="（木）";
	cWeek[5]="（金）";
	cWeek[6]="（土）";
	todayText=toYear+'年 '+toMonth+'月 '+toDay+'日'+cWeek[toDate.getDay()];
	m_Sun = "日";
	m_Mon = "月";
	m_Tus = "火";
	m_Wed = "水";
	m_Thu = "木";
	m_Fri = "金";
	m_Sat = "土";
	m_Coment = "前月／翌月のカレンダーは，▲/▼を押してください．";

var dYY = new Array();
var dMM = new Array();
var dDD = new Array();
var dDDlast = new Array();
var dfgColor = new Array();
var dbgColor = new Array();
// 年(dYY)、月(dMM)が0は週カレンダー　日(dDD)に特定の曜日の色を指定します。　0=日曜,1=月曜,..,6=土曜
dYY[0]=0;dMM[0]=0;dDD[0]=6;dDDlast[0]=0;dfgColor[0]="#8A2BE2";dbgColor[0]="";
// 日(dDD)が負の場合は、月の特定曜日になります。-NW N:第ｎ曜日(N=1,2,3,4,5)、W:曜日(W=0,1,2,3,4,5,6) 0=日曜,1=月曜,..,6=土曜。 例：-21 は第2月曜日
// 年共通カレンダー　毎年の特定月日（祝日など）の色を指定します。月日は昇順に記述してください。
dYY[1]=0;dMM[1]=1;dDD[1]=1;dDDlast[1]=0;dfgColor[1]="red";dbgColor[1]="";
dYY[2]=0;dMM[2]=1;dDD[2]=-21;dDDlast[2]=0;dfgColor[2]="red";dbgColor[2]="";
dYY[3]=0;dMM[3]=2;dDD[3]=11;dDDlast[3]=0;dfgColor[3]="red";dbgColor[3]="";
dYY[4]=0;dMM[4]=3;dDD[4]=21;dDDlast[4]=0;dfgColor[4]="red";dbgColor[4]="";
dYY[5]=0;dMM[5]=4;dDD[5]=29;dDDlast[5]=0;dfgColor[5]="red";dbgColor[5]="";
dYY[6]=0;dMM[6]=5;dDD[6]=3;dDDlast[6]=0;dfgColor[6]="red";dbgColor[6]="";
dYY[7]=0;dMM[7]=5;dDD[7]=4;dDDlast[7]=0;dfgColor[7]="red";dbgColor[7]="";
dYY[8]=0;dMM[8]=5;dDD[8]=5;dDDlast[8]=0;dfgColor[8]="red";dbgColor[8]="";
dYY[9]=0;dMM[9]=7;dDD[9]=20;dDDlast[9]=0;dfgColor[9]="red";dbgColor[9]="";
dYY[10]=0;dMM[10]=9;dDD[10]=15;dDDlast[10]=0;dfgColor[10]="red";dbgColor[10]="";
dYY[11]=0;dMM[11]=9;dDD[11]=23;dDDlast[11]=0;dfgColor[11]="red";dbgColor[11]="";
dYY[12]=0;dMM[12]=10;dDD[12]=-21;dDDlast[12]=0;dfgColor[12]="red";dbgColor[12]="";
dYY[13]=0;dMM[13]=11;dDD[13]=3;dDDlast[13]=0;dfgColor[13]="red";dbgColor[13]="";
dYY[14]=0;dMM[14]=11;dDD[14]=22;dDDlast[14]=0;dfgColor[14]="red";dbgColor[14]="";
dYY[15]=0;dMM[15]=12;dDD[15]=23;dDDlast[15]=0;dfgColor[15]="red";dbgColor[15]="";
//　年のカレンダー　特定年月日の色を指定します。年月日は昇順に記述してください。
// 例　dYY[16]=1998;dMM[16]=12;dDD[16]=31;dDDlast[16]=0;dfgColor[16]="green";dbgColor[16]="";
// NoOfdItems に指定した配列の数を指定してください。
NoOfdItems=16;
var docText=""

// Calender
function MakeCalender(){
	
	var oDoc = frames["calender"].document;
	oDoc.close();
	
	var Month = sDate.getMonth()+1;
	sDate.setMonth(Month - 1);
	sDate.setDate(31);
	sDate.setMonth(Month - 1);
	var DD=sDate.getDate();
	sDate.setDate(1);
	sDate.setMonth(Month - 1);
	var WeekDay = sDate.getDay();
	var Year = sDate.getYear()
	if( Year < 2000){	Year += 1900 };
	
	
	for(Kx=0; Kx < NoOfdItems; Kx++){
		if( dYY[Kx] > 0 || dMM[Kx] > 0 ){	break;}
	}
	var endW = Kx;
	var bgnM = Kx;
	for(Kx=bgnM; Kx < NoOfdItems; Kx++){
		if( dYY[Kx] == 0 && dMM[Kx] < Month ){	bgnM++;}
		if( dYY[Kx] > 0 || dMM[Kx] > Month ){	break;}
	}
	var endM = Kx;
	var bgnY = Kx;
	for(Kx=bgnY; Kx < NoOfdItems; Kx++){
		if( dYY[Kx] < Year ){	bgnY++;}
		if( dYY[Kx] == Year && dMM[Kx] < Month ){	bgnY++;}
		if( dYY[Kx] >= Year && dMM[Kx] > Month ){	break;}
	}
	var endY = Kx;

	docText ='<HTML>';
	docText+='<HEAD><link rel="stylesheet" type="text/css" href="stCal.css"></HEAD>';
	docText+='<BODY>';
	docText+='<CENTER><TABLE><TR><TD ALIGN=CENTER>';
	docText+='<A HREF="" onclick="return parent.setText(' + toYear + ',' + toMonth + ',' + toDay + ')">';
	docText+= todayText+'</TD></TR></TABLE>';
	docText+='<TABLE BORDER=4 CELLPADDING=0 CELLSPACING=0 BGCOLOR=#ffffff width="90%"><TR><TD>' ;
	docText+='<CENTER><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0 >' ;
	docText+='<TR>';
	docText+='<TD ALIGN=CENTER><A HREF="javascript:parent.mBack()">▲</A></TD>';
	docText+='<TD COLSPAN=5 ALIGN=CENTER>';
	
	docText+=Year+'年 '+Month+'月</TD>';
	
	docText+='<TD ALIGN=CENTER><A HREF="javascript:parent.mForward()">▼</A></TD>';
	docText+='</TR><TR BGCOLOR="#cccccc">';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:FF0000">'+m_Sun+'</span></TD>\n';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:000000">'+m_Mon+'</span></TD>\n';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:000000">'+m_Tus+'</span></TD>\n';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:000000">'+m_Wed+'</span></TD>\n';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:000000">'+m_Thu+'</span></TD>\n';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:000000">'+m_Fri+'</span></TD>\n';
	docText+='<TD ALIGN=CENTER WIDTH=24 ><span style="color:8A2BE2">'+m_Sat+'</span></TD>\n';
	docText+='</TR><TR><TD COLSPAN=7></TD></TR><TR>';

	oDoc.writeln(docText);docText="";
	Days=31;
	if(DD < 31){	Days= 31 - DD;}
	DD= Days+WeekDay;
	var Wx = Math.ceil(DD/7);
	for(Jx=0; Jx < WeekDay; Jx++){docText+='<TD WIDTH=24></TD>';}
	Jx=WeekDay
	var DD=1;
	for(Ix=0; Ix < Wx; Ix++){
		for(Jx; Jx<7; Jx++){
			var MW= -((Ix+1)*10 + Jx)
			if(Jx < WeekDay)	MW+=10;
			
			if( Jx == 0 ){
				var fColor = "#FF0000";
			}else{	
				fColor = "#000000";
			}
			
			var bColor="";
			
			for(Kx=0; Kx < endW; Kx++){
				if( Jx >dDD[Kx] && Jx <= dDDlast[Kx] ){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					bColor = dbgColor[Kx];
				}
				
				if( Jx == dDD[Kx] || MW == dDD[Kx] ){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					
					bColor = dbgColor[Kx];
					break;
				}
			}
			for(Kx=bgnM; Kx < endM; Kx++){
				if( Jx == 1 && DD == dDD[Kx] +1 ){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					break;
				}
				
				if( DD >dDD[Kx] && DD <= dDDlast[Kx]  ){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					bColor = dbgColor[Kx];
				}
				if( DD == dDD[Kx] || MW == dDD[Kx] ){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					bColor = dbgColor[Kx];
					break;
				}
			}
			for(Kx=bgnY; Kx < endY; Kx++){
				if( dMM[Kx] == Month &&  DD >dDD[Kx] && DD <= dDDlast[Kx]  ){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					bColor = dbgColor[Kx];
				}
				
				if( dMM[Kx] == Month && ( DD == dDD[Kx] || MW == dDD[Kx] )){
					if( dfgColor[Kx] > ""){
						fColor = dfgColor[Kx];
					}
					bColor = dbgColor[Kx];
					
					break;
				}
			}
			docText+='<TD WIDTH=20 ALIGN=RIGHT';
			
			if( bColor != ""){
				docText+=' BGCOLOR="'+bColor+'" >';
			}else{
				docText+='>';
			}
			
// setText関数の処理変更に伴う変更
//			docText+='<A HREF="javascript:parent.setText(' + Year + ',' + Month + ',' + DD + ')">'; 
			docText+='<A HREF="" onclick="return parent.setText(' + Year + ',' + Month + ',' + DD + ')">';
			
			docText+='<span style="color:'+fColor+'">';

			strDD = new String(DD)

			if(strDD.length == 1){
				strDD = "&nbsp" + strDD;
			}

			if(Year==thisYear && Month==thisMonth && DD==thisDate){
				docText+='<span class="today">'+strDD+'</span>';
			}else{
				docText+=strDD;
			}
			
			docText+='</span></A></TD>\n';
			DD++;
			
			if(DD > Days)break;
		}
		Jx=0;
		docText+='</TR><TR>'
		oDoc.writeln(docText);docText="";
	}
	docText+='</TR></TABLE></CENTER>'
	docText+='</TD></TR></TABLE>'
	docText+=m_Coment+'</CENTER>'
	docText+="<BR>"
	docText+="</BODY></HTML>";
	
	oDoc.writeln(docText);docText="";
	oDoc.close();
}
function mBack(){
	var YY=sDate.getYear();
	if( YY < 2000 ){ YY+=1900;}
	var MM=sDate.getMonth();
	if(MM ==0){
		YY--;
		sDate.setYear(YY);
		MM=12;
	}
	sDate.setMonth(MM-1);
	sDate.getDate()
	sDate.setMonth(MM-1);
	MakeCalender();
}
function mForward(){
	var YY=sDate.getYear() ;
	if( YY < 2000 ){ YY+=1900;}
	var MM=sDate.getMonth();
	if(MM ==11){
		YY++;
		sDate.setYear(YY);
		MM=-1;
	}
	sDate.setMonth(MM+1);
	sDate.getDate()
	sDate.setMonth(MM+1);
	MakeCalender();
}
function setText(yy,mm,dd){

    // 今のプログラムをwURLにいれる
    var wURL = opener.location.href;
        wURL.match(/.asp/);
        wURL = RegExp.leftContext;
	
	// 一致する場合のみ作動
	if (wDefURL == wURL) {
		if(mm < 10 ) mm = "0" + mm
		if(dd < 10 ) dd = "0" + dd
		
		parent.opener.document.all[hid_txtname.value].value = yy + "/" + mm + "/" + dd;
		self.close();
	} else {
		alert("ページを移動したため、選択できません。");
		return false;
	}
}
function checkDateFormat(str){
	cd = new Date(str);
	alert("cd = " +cd);
	if(cd=="NaN"){
		if(str==""){
			return true;
		}else{
			return false;
		}
	}else{
		if(str.match(/[^0123456789/]/i)!=null) return false;
		return true;
	}
}
//-->
</script>

</HEAD>

<FRAMESET ROWS="*" onLoad="MakeCalender()" border=0>
	<FRAME SRC="javascript:'<html><body BGCOLOR=#FFFFFF></body></html>'" NAME="calender" border=0 SCROLLING="NO">
</FRAMESET>

</BODY>
</HTML>
