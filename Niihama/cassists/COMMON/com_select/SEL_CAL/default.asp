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
<title>�J�����_�[</title>
</head>
<script language="JavaScript"><!--
/* use free!! , TAB = 4 */
/* ���ꉞ�A���O�����Ă͂���܂����A���쌠�͕������܂��B�䎩�R�ɂ��g�����������B */
/* ���g�����͏����Ă����ł��B(��) */

// by P,H,M (����)
// modified by Mugi
// original=http://www.wakusei.ne.jp/~yuuki/web/test/00456.html
//
// �ȉ��A�� �����ҏ���(��)
//
// thank's �������܂���		��Ԃ͂��߂̒p����������
// thank's ���ɂǂ��H����	��Ԃ͂��߂̒p����������
// thank's �������񂳂�		�l�X�P�̌�
// thank's JYA^2����		�[�N����
// thank's Three-eye����	�x�������S�ʁA���̑��F�X(^^)
// thank's NobMochi����		����m�F
// thank's Joga����			�u�t���̓��v�u�H���̓��v

function mz_calendar(mondiff){
	var		j,z;
	var		today,display;
	var		year,mon,date,day;
	var		nday;
	var		topday;	// ���̍ŏ��̗j��
	var		endd;	// ���̍ŏI�̓��ɂ�
	var		wk;		// ��?�T
	var		ln;		// �J�����_�[�\�����?�i
	var		furikae	// �U�֋x�������p�t���O�B1�Ȃ玟�̓��͐U�֋x��
	var		VernalEquinoxDayd  ,VernalEquinoxDay;	// �t���̓�
	var		AutumnalEquinoxDayd,AutumnalEquinoxDay;	// �H���̓�

	/***** �\�����镶���̑�����` *****/
	var		COLOR  =   7;		// 0000.0111 MASK 8�F
	var		ATTRIB = 248;		// 1111.1000 MASK 6����
	var		NORM   =   0;		// 000	�W���F(�F�w�������̕����F)
	var		RED    =   1;		// 001	�ԐF
	var		BLUE   =   2;		// 010	�F
//	var		       =   3;		// 011
//	var		       =   4;		// 011
//	var		       =   5;		// 011
//	var		       =   6;		// 011
//	var		       =   7;		// 011
	var		UL     =   8;		// ����
	var		SO     =  16;		// �����
	var		IT     =  32;		// �C�^���b�N
//	var		       =  64;		// �\��
//	var		       = 128;		// �\��
//	var		       = 256;		// �\��
	now = new Date("<%=m_stoday%>");
	
	
	year = now.getYear();	if( year<1900 ) year+=1900;
	
	mon  = now.getMonth()+1+mondiff;
	while( 12<mon ){	year++;		mon-=12;	}
	while( mon< 1 ){	year-=1;	mon+=12;	}
	date = now.getDate();
	
	nday  = now.getDay();
	if( nday==0 )	day = "��";
	if( nday==1 )	day = "��";
	if( nday==2 )	day = "��";
	if( nday==3 )	day = "��";
	if( nday==4 )	day = "��";
	if( nday==5 )	day = "��";
	if( nday==6 )	day = "�y";

	// �����̍ŏ��̗j���� topday �ɓ���B
	topd = new Date(year,mon-1,1);	// ������ 1���̏��� topd �ɓ���B
//	topd = new Date("<%=m_iyear%>",<%=m_imonth%>,1);	// ������ 1���̏��� topd �ɓ���B
	topday = topd.getDay();
	// �����̍ŏI���� endday �ɓ���B ��NC4 �ł͓��삵�܂���ł����BNN2,3,IE3,4,5�Ȃ� OK �Ȃ̂ɂ�(;_;)
	//	endd = new Date(year,mon,0);	// �����̍ŏI���̏��� endd �ɓ���B
	//	endday = endd.getDate();		// endday = �����̍ŏI���t

	if( mon==2 ){
		if( (year%400)==0 )			endday = 29;	/* 400 �Ŋ��؂��N��2/29�܂� */
		else if( (year%100)==0 )	endday = 28;	/* 100 �Ŋ��؂��N��2/28�܂� */
		else if( (year%  4)==0 )	endday = 29;	/*   4 �Ŋ��؂��N��2/29�܂� */
		else						endday = 28;	/* ���̑��̔N�� 2/28 �܂� */	
	}																			
	else if( (mon==4)||(mon==6)||(mon==9)||(mon==11) )	endday = 30;				
	else	endday = 31;															

	/***** ���N�ɂ���ĕω�����u�t���̓��v�u�H���̓��v���ȈՌv�Z�ɂ�苁�߂� *****/
	// �Q�l�Fhttp://www.top.or.jp/~cpop/syunbun.htm
	// 1900�`2099�N�͈̔͌���i2100�N�ɂȂ�����A�ӔC�������ďC�����Ă��������B(^^; �j

	// �����ȓ��t�́g�O�̔N��2��1���t���̊���Ō��������h���ƂɂȂ��Ă��܂��B
	// �ȈՌv�Z�ɂ���ċ��߂��l������Ȃ��N������������A���̔N�ɂ��āA�K�X�A��O�������ĉ������B

	VernalEquinoxDayd   = Math.floor(0.24242*year - Math.floor(year/4) + 35.84);	// year�N�̎ҕ��̓�
	AutumnalEquinoxDayd = Math.floor(0.24204*year - Math.floor(year/4) + 39.01);	// year�N�̏H���̓�

	// ����O�����̗�
	//	if( year==20XX ){	// 20XX�N�Ȃ�
	//		VernalEquinoxDayd   = 21;	// �t���̓� = 21��(3��)
	//		AutumnalEquinoxDayd = 23;	// �H���̓� = 23��(9��)
	//	}
	//	else{	// ���̑��̔N�Ȃ�ȈՌv�Z�l�̂܂܂� OK
	//		VernalEquinoxDayd   = Math.floor(0.24242*year - Math.floor(year/4) + 35.84);	// year�N�̎ҕ��̓�
	//		AutumnalEquinoxDayd = Math.floor(0.24204*year - Math.floor(year/4) + 39.01);	// year�N�̏H���̓�
	//	}

	VernalEquinoxDay   = 3+VernalEquinoxDayd  /100;		// year�N�̏t���̓�
	AutumnalEquinoxDay = 9+AutumnalEquinoxDayd/100;		// year�N�̏H���̓�
	/***** ���u�t���̓��v�u�H���̓��v�����܂Ł� *****/

	/* �ȉ��u�L���L�������₡�I�I�v */
		pre = <%=m_iMondiff%> -1;
		nxt = <%=m_iMondiff%> +1;
		document.write("<pre><tt>");
		document.write("<a href='javascript:nextcal("+pre+");' onclick='nextcal("+pre+")'><<</a>  ");
//		if( mon==3 )		document.write( "   "    + year +"/"+ mon+" �t��" + VernalEquinoxDayd   + "��" );
//		else if( mon==9 )	document.write( "   "    + year +"/"+ mon+" �H��" + AutumnalEquinoxDayd + "��" );
//		else{
							if( mon<10  )		document.write(" ");	// �X�y�[�X�̒���
							document.write( "   " + year +"/"+ mon +"   " );
//		}
		document.writeln("  <a href='javascript:nextcal("+nxt+");' onclick='nextcal("+nxt+")'>>></a>");
//		document.writeln("<table border='0' width='100%'><tr>");
//		document.writeln("<td align='left'><a href='#' onclick='nextcal("+pre+")'><<</a></td>");
//		document.writeln("<td align='right'><a href='#' onclick='nextcal("+nxt+")'>>></a></td>");
//		document.writeln("</tr></table>");

	document.writeln("");
	document.write('<font color="#FF0000">��<\/font> �� �� �� �� �� <font color="#0000FF">�y<\/font><br>');
	for( j=0 ;j<topday;j++)	document.write("   ");

	z=j+1;	// z  = �j��(1=��/2=��...7=�y)
	ln = 1;

	for( j=1 ; j<=endday ; j++,z++ ){		// ���P�����\���̃��[�v
		wk = 1;					// ��1�T
		if(  8<=j )	wk=2;		// ��2�T
		if( 15<=j )	wk=3;		// ��3�T
		if( 22<=j )	wk=4;		// ��4�T
		if( 29<=j )	wk=5;		// ��5�T

		// �ȍ~�� j�F���t / z�F�j��(1=��/2=��...7=�y) / wk�F��?�T / ln�F�J�����_�[�\�����?�i��
		//        year�F�\������N / mon�F�\�����錎 / date�F�����̓��ɂ�

		today = eval(mon+"+"+(j/100));		// today = ��.��

		/***** ��������x���\������ *****/
		display = 0;	// ���t�̐����̕����F�E�������������Z�b�g���Ă����B

		// �U�֋x������(�������Ɂu�����̋x���v�͖����Ƃ����O��ō���Ă��܂��B)
		if( furikae )	display = RED|UL;	// �O�����u�����̏j���v�Łu���j���v�������灨���x�݁B
											// �U�֋x�����ԐF������
		// �������̏j��
		if( today== 1.01 )	display = RED;	// ����
		if( today== 2.11 )	display = RED;	// �����L�O�̓�
		if( today== 4.29 )	display = RED;	// �݂ǂ�̓�
		if( today== 5.03 )	display = RED;	// ���@�L�O��
		if( today== 5.05 )	display = RED;	// ���ǂ��̓�
		if( today== 7.20 )	display = RED;	// �C�̓�
		if( today== 9.15 )	display = RED;	// �h�V�̓�
		if( today==11.03 )	display = RED;	// �����̓�
		if( today==11.23 )	display = RED;	// �ΘJ���ӂ̓�
		if( today==12.23 )	display = RED;	// �V�c�a����
		if( today == VernalEquinoxDay   )	display = RED;	// �t���̓�
		if( today == AutumnalEquinoxDay )	display = RED;	// �H���̓�

		// �u���l�̓��v�u�̈�̓��v����
		if( 2000>year ){		// 2000�N���O
			if( today== 1.15 )	display = RED;	// ���l�̓�
			if( today==10.10 )	display = RED;	// �̈�̓�
		}
		else{					// 2000�N����
			if( ( mon==1  )&&(z==2)&&(wk==2) ){	display = RED;	}	// ���l�̓�( 1����2�T�̌��j��)
			if( ( mon==10 )&&(z==2)&&(wk==2) ){	display = RED;	}	// �̈�̓�(10����2�T�̌��j��)
		}

		// �U�֋x������
		// �������Ɂu�����̋x���v�͖����i1�����U�֋x���ɂȂ邱�Ƃ��Ȃ��j�Ƃ����O��ō���Ă��܂��B
		if( ( display )&&( z<=1 ) )	furikae = 1;	// �������u�����̏j���v�Łu���j���v�������玟�̓��͋x��
		else						furikae = 0;

		if( today== 5.04 )	display = RED;	// �u���@�L�O���v�Ɓu���ǂ��̓��v�ɋ��܂�Ă��邩��x��
		// �g�u�����̏j���v�ɋ��܂ꂽ1���͋x���Ƃ���h�ɊY������̂͌��݂��̓������B(����)


		// ���[�U�[�J�X�^�}�C�Y�̈�
		if( z<=1 )			display = 1;	// ���j��
//		if( z==7 )			display = 2;	// �y�j��		�y�j����ŕ\����������
		if( z==7 )			display = 1;	// �y�j��		�y�j����Ԃŕ\����������

		if( year==1998 ){						// 1998�N�x���f�[�^
			if( today== 1.02 )	display = RED;		// �x��
			if( today== 1.17 )	display = BLUE|UL;	// �y�j��		�F������
			if( today== 2.14 )	display = BLUE|UL;	// �y�j��		�F������
			if( today== 4.30 )	display = RED|UL;	// �U��			�ԐF������
			if( today== 5.01 )	display = RED;		// ??????
			if( today== 7.27 )	display = RED;		// �ċx��
			if( today== 7.28 )	display = RED;		// �ċx��
			if( today== 7.29 )	display = RED;		// �ċx��
			if( today== 7.30 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 7.31 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 8.14 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 9.19 )	display = BLUE|UL;	// �y�j��		�F������
			if( today==11.07 )	display = BLUE|UL;	// �y�j��		�F������
			if( today==11.28 )	display = BLUE|UL;	// �y�j��		�F������
			if( today==12.29 )	display = RED;		// �x��
			if( today==12.30 )	display = RED;		// �x��
			if( today==12.31 )	display = RED;		// �x��
		}
		if( year==1999 ){						// 1999�N�x���f�[�^
			if( today== 1.09 )	display = BLUE;		// �y�j��
			if( today== 1.04 )	display = RED;		// �x��
			if( today== 3.22 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 4.30 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 7.20 )	display = NORM|UL;	// �U��ւ�		�W���F������
			if( today== 7.26 )	display = RED;		// �ċx��
			if( today== 7.27 )	display = RED;		// �ċx��
			if( today== 7.28 )	display = RED;		// �ċx��
			if( today== 7.29 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 7.30 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 8.13 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 8.16 )	display = RED|UL;	// �U��ւ�		�ԐF������
			if( today== 9.18 )	display = BLUE|UL;	// �y�j��		�F������
			if( today==10.16 )	display = BLUE|UL;	// �y�j��		�F������
			if( today==11.06 )	display = BLUE|UL;	// �y�j��		�F������
			if( today==12.29 )	display = RED;		// �x��
			if( today==12.30 )	display = RED;		// �x��
			if( today==12.31 )	display = RED;		// �x��
		}
		if( year==2000 ){						// 2000�N�x���f�[�^
			if( today== 1.03 )	display = RED;		// �x��
			if( today== 1.04 )	display = RED;		// �x��
			if( today== 1.15 )	display = BLUE|UL;	// �y�j��			�F������
		}


		if( display & UL )				document.write('<u>');							// ����
		if( display & SO )				document.write('<s>');							// �����
		if( display & IT )				document.write('<i>');							// �C�^���b�N
		if( (display&COLOR)==RED  )		document.write('<font color="#FF0000">');		// ��
		if( (display&COLOR)==BLUE )		document.write('<font color="#0000FF">');		// ��
		if( (mondiff==0)&&(j==date) )   document.write("<FONT style='background:#55FF55'>");	// ����

		if( j<10 )	document.write(" ");		// ���t��1���Ȃ�X�y�[�X��1����
		document.write('<a href=# class=datelink onclick="inputdate('+year+','+mon+','+j+');return false">'+j+'</\a>');						// ���t���������� "��(- -;)

		if( (mondiff==0)&&(j==date) )	document.write("<\/FONT>");						// ����
		if( (display&COLOR)==BLUE )		document.write('<\/font>');						// ��
		if( (display&COLOR)==RED  )		document.write('<\/font>');						// ��
		if( display & IT )				document.write('<\/i>');						// �C�^���b�N
		if( display & SO )				document.write('<\/s>');						// �����
		if( display & UL )				document.write('<\/u>');						// ����

		if( 6<z ){	z=0;	ln++;	document.write("<br>");	}	// �T�I��聨���s
		else						document.write(" ");		// ���t�Ԃ̃X�y�[�X
	}
				document.write("<\/tt><\/pre>");
}

function mz_clock(){
	var hour,mini,sec;

	now = new Date();
	hour = now.getHours();		if( hour<10 )	hour = "0"+hour;
	mini = now.getMinutes();	if( mini<10 )	mini = "0"+mini;
	sec  = now.getSeconds();	if( sec <10 )	sec  = "0"+sec ;

//  form �� name�� ��input �� name
	document.clock.time.value = hour + ":" + mini + "," + sec ;
	setTimeout("mz_clock()",500);	// ������s�� 500/1000(=0.5)�b��
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
<input type="button" class="button" value="����" onClick="window.close()">
</center>
</body>
</html>
