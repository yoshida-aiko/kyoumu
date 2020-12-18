<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 部活動部員一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0360/web0360_main.asp
' 機      能: 部員を表示
'-------------------------------------------------------------------------
' 引      数:   txtClubCd		:部活CD
'               cboGakunenCd	:学年
'               cboClassCd		:クラスNO
'               txtTyuClubCd	:中学校部活CD
'
' 引      渡:   txtMode			:処理モード
'               txtClubCd		:部活CD
'               GAKUSEI_NO		:学生NO
'               cboGakunenCd	:学年
'               cboClassCd		:クラスNO
'               txtTyuClubCd	:中学校部活CD
' 説      明:
'           ■初期表示
'               空白ページを表示
'           ■表示ボタンが押された場合
'               ・検索条件にかなった生徒一覧を表示する
'               ・部活が二つとも埋まっている生徒の入部登録は不可とする(選択チェックボックスを表示しない)
'               ・すでに登録対象部活に入部している生徒の入部登録は不可とする(選択チェックボックスを表示しない)
'-------------------------------------------------------------------------
' 作      成: 2001/08/22 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public m_iSyoriNen			'//年度
	Public m_iKyokanCd			'//教官ｺｰﾄﾞ
	Public m_sClubCd			'//クラブCD
	Public m_iGakunen           '//学年
	Public m_iClassNo           '//クラスNO
	Public m_sTyuClubCd			'//中学校クラブCD

    'ﾚｺｰﾄﾞセット
	Public m_Rs					'//部員一覧ﾚｺｰﾄﾞｾｯﾄ
	Public m_iRsCnt				'//ﾚｺｰﾄﾞカウント

	'エラー系
	Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Dim w_iRet  			'// 戻り値

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="部活動部員一覧"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			Call gs_SetErrMsg("データベースとの接続に失敗しました。")
			Exit Do
		End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'//値の初期化
		Call s_ClearParam()

		'//変数セット
		Call s_SetParam()

'//デバッグ
'Call s_DebugPrint()

		'//生徒一覧の取得
		w_iRet = f_GetSeitoData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'// ページを表示
		Call showPage()
		Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(m_Rs)

	'// 終了処理
	Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

	m_iSyoriNen  = ""
	m_iKyokanCd  = ""
	m_sClubCd  	= ""
	m_iGakunen   = ""
	m_iClassNo   = ""
	m_sTyuClubCd = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

	m_iSyoriNen  = Session("NENDO")
	m_iKyokanCd  = Session("KYOKAN_CD")
	m_sClubCd    = Request("txtClubCd")
	m_iGakunen   = Request("cboGakunenCd")	'//学年
	m_iClassNo   = gf_cboNull(Request("cboClassCd"))	'//クラス
	m_sTyuClubCd = replace(Request("txtTyuClubCd"),"@@@","")	'//中学校クラブCD
	Session("HyoujiNendo") = m_iSyoriNen
End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
	response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
	response.write "m_sClubCd    = " & m_sClubCd    & "<br>"
	response.write "m_iGakunen   = " & m_iGakunen   & "<br>"
	response.write "m_iClassNo   = " & m_iClassNo   & "<br>"
	response.write "m_sTyuClubCd = " & m_sTyuClubCd & "<br>"

End Sub

'********************************************************************************
'*  [機能]  生徒一覧情報取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSeitoData()

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetSeitoData = 1

	Do

		'//生徒一覧
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_SIMEI"
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_TYU_CLUB"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO"
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_GAKUNEN=" & m_iGakunen
		
		if m_iClassNo <> "" then
				w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_CLASS=" & m_iClassNo
		End if
		If m_sTyuClubCd <> "" Then
			w_sSql = w_sSql & vbCrLf & "  AND  T11_GAKUSEKI.T11_TYU_CLUB='" & m_sTyuClubCd & "'"
		End If

		w_sSql = w_sSql & vbCrLf & " ORDER BY "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN, T13_GAKU_NEN.T13_CLASS, T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write w_sSQL & "<br>"
		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_GetSeitoData = 99
			Exit Do
		End If

		'//ﾚｺｰﾄﾞカウント取得
		'//件数を取得
		m_iRsCnt = 0
		If m_Rs.EOF = False Then
			m_iRsCnt = gf_GetRsCount(m_Rs)
		End If

		'//正常終了
		f_GetSeitoData = 0
		Exit Do
	Loop


End Function

'********************************************************************************
'*  [機能]  クラス情報を取得
'*  [引数]  p_iGakuNen:学年,p_iClassNo:クラスNO
'*  [戻値]  f_GetClassName:クラス名
'*  [説明]  
'********************************************************************************
Function f_GetClassName(p_iGakuNen,p_iClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassName = ""
	w_sClassName = ""

	Do
		'クラスマスタよりデータを取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSMEI"
		w_sSql = w_sSql & vbCrLf & "  ,M05_CLASS.M05_GAKKA_CD"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN= " & p_iGakuNen
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO= "   & p_iClassNo

'response.write w_sSQL & "<br>"

		'//データ取得
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			Exit Do
		End If

		If rs.EOF = False Then
			w_sClassName = rs("M05_CLASSMEI")
			'w_sGakkaCd = rs("M05_GAKKA_CD")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetClassName = w_sClassName

	'//ﾚｺｰﾄﾞCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  部活名を取得する
'*  [引数]  p_sClubCd:部活CD
'*  [戻値]  f_GetClubName：部活名
'*  [説明]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//部活CDが空の時
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//部活動情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//部活名
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetClubName = w_sClubName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

%>


	<html>
	<head>
	<title>部活動部員一覧</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//  [機能]  ページロード時処理
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function window_onload() {
		<%If m_Rs.EOF = false Then%>
//			parent.topFrame.document.frm.btnShow.disabled=true;
//			parent.topFrame.document.frm.cboGakunenCd.disabled=true;
//			parent.topFrame.document.frm.cboClassCd.disabled=true;
//			parent.topFrame.document.frm.txtTyuClubCd.disabled=true;
		<%End If%>
	}
	//************************************************************
	//  [機能]  登録ボタンが押されたとき
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_Touroku(){

		// 入力値のﾁｪｯｸ
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

		if (!confirm("登録してもよろしいですか？")) {
		return ;
		}

		//リスト情報をsubmit
		document.frm.txtMode.value = "INSERT";
		document.frm.target = "main";
		document.frm.action = "./web0360_edt.asp"
		document.frm.submit();
		return;
	}

	//************************************************************
	//  [機能]  キャンセルボタンが押されたとき
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_Cancel(){

		//上画面使用可とする
//		parent.topFrame.document.frm.btnShow.disabled=false;
//		parent.topFrame.document.frm.cboGakunenCd.disabled=false;
//		parent.topFrame.document.frm.cboClassCd.disabled=false;
//		parent.topFrame.document.frm.txtTyuClubCd.disabled=false;

		//空白ページを表示
		parent.main.location.href="default3.asp?txtClubCd=<%=m_sClubCd%>"

	}

	//************************************************************
	//  [機能]  戻るボタンが押されたとき
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_Back(){
		//キャンセル時、初期画面に戻る
		//上フレーム再表示
		parent.topFrame.location.href="./web0360_top.asp?txtClubCd=<%=m_sClubCd%>"
		//下フレーム再表示
		parent.main.location.href="./web0360_main.asp?txtClubCd=<%=m_sClubCd%>"

	}

    //************************************************************
    //  [機能]  チェック欄がチェックされているか
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //************************************************************
    function f_CheckData() {

		//チェック欄数を取得
		var iMax = document.frm.chkMax.value
		if (iMax==0){
			//alert("No Avairable")
			return 1;
		}

		if(iMax==1){
			if(document.frm.GAKUSEI_NO.checked==false){
				alert("登録する生徒が選択されていません")
				return 1;
			}
		}else{

			var i
			var w_bCheck = 1
			for (i = 0; i < iMax; i++) {
				if(document.frm.GAKUSEI_NO[i].checked==true){
					w_bCheck = 0
					break;
				}
			};

			if(w_bCheck == 1){
				alert("登録する生徒が選択されていません")
				return 1;
			};
		};

        return 0;
    }

    //************************************************************
    //  [機能]  詳細ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_detail(pGAKUSEI_NO){

			url = "/cassist/gak/gak0310/kojin.asp?hidGAKUSEI_NO=" + pGAKUSEI_NO;
			w   = 700;
			h   = 630;

			wn  = "SubWindow";
			opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

//		document.frm.hidGAKUSEI_NO.value = pGAKUSEI_NO;
//		document.forms[0].submit();
    }
	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<center>
	<form name="frm" method="post">

	<%
	Do
		'//データなしの場合
		If m_Rs.EOF Then%>
			<br><br>
			<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
			<br><br>
			<!--
			<input class="button" type="button" onclick="javascript:f_Cancel();" value="キャンセル">
			&nbsp;&nbsp;&nbsp;&nbsp;
			-->
			<input class="button" type="button" onclick="javascript:f_Back();" value="　戻　る　">

			<%Exit Do
		End If
	%>

		<br>

		<table>
			<tr>
			<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="　登　録　"></td>
			<td ><input class="button" type="button" onclick="javascript:f_Cancel();"  value="キャンセル"></td>
			<td ><input class="button" type="button" onclick="javascript:f_Back();"    value="　戻　る　"></td>
			</tr>
		</table>

		<span class="msg">＊入部登録する生徒を選択し、「登録」ボタンを押してください。</span><br>
		<span class="msg">＊すでに入部している生徒は登録できません。</span>

		<!--リスト部-->
		<table >
			<tr><td valign="top">

				<table class=hyo border="1" bgcolor="#FFFFFF">

					<!--ヘッダ-->
					<tr>
						<th nowrap class="header" width="15"  align="center">選<br>択</th>
						<th nowrap class="header" width="40"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
						<th nowrap class="header" width="120" align="center">氏名</th>
						<th nowrap class="header" width="100" align="center">現部活</th>
<!--
						<th nowrap class="header" width="100" align="center">中学部活</th>
-->
					</tr>

			<%
			j = 0
			w_iCnt = INT(m_iRsCnt/2 + 0.9)
			Do until m_Rs.EOF 

				'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
				Call gs_cellPtn(w_Class)
				i = i + 1
				%>
					<tr>
						<td nowrap class="<%=w_Class%>" width="15"  align="center" rowspan="2">

						<%
                        '//すでにクラブ1とクラブ2にデータがある生徒は、クラブの新規登録は不可とする
						'If gf_SetNull2String(m_Rs("T13_CLUB_1")) <> "" AND gf_SetNull2String(m_Rs("T13_CLUB_2")) <> "" Then
                        '//すでにクラブ1とクラブ2にデータがあり両方とも入部中の生徒は、クラブの新規登録は不可とする 2001/12/11 伊藤
						If gf_SetNull2String(m_Rs("T13_CLUB_1")) <> "" AND gf_SetNull2String(m_Rs("T13_CLUB_2")) <> "" AND gf_SetNull2String(m_Rs("T13_CLUB_1_FLG")) = "1" AND gf_SetNull2String(m_Rs("T13_CLUB_2_FLG")) = "1" Then
                        %>
							<br>
						<%Else%>

							<%
							'//すでに登録対象クラブに所属している生徒は、更新不可とする
							'If (gf_SetNull2String(m_Rs("T13_CLUB_1")) = m_sClubCd) Or (gf_SetNull2String(m_Rs("T13_CLUB_2")) = m_sClubCd) Then
							If (gf_SetNull2String(m_Rs("T13_CLUB_1")) = m_sClubCd And gf_SetNull2String(m_Rs("T13_CLUB_1_FLG")) = "1") Or (gf_SetNull2String(m_Rs("T13_CLUB_2")) = m_sClubCd And gf_SetNull2String(m_Rs("T13_CLUB_2_FLG")) = "1") Then
							%>
								<br>
							<%Else
								j = j + 1
								%>
								<input type="checkbox" name="GAKUSEI_NO" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<%End If%>

						<%End If%>

						</td>
						<td nowrap class="<%=w_Class%>" width="40"  align="left"   rowspan="2"><%=m_Rs("T13_GAKUSEKI_NO")%><br></td>
						<td nowrap class="<%=w_Class%>" width="120" align="left"   rowspan="2"><a href="#" onClick="f_detail(<%=m_Rs("T13_GAKUSEI_NO")%>)"><%=m_Rs("T11_SIMEI")%></a><br></td>
						<td class="<%=w_Class%>" width="100" align="left" >
							
							<%
							'入部中なら表示する
							If gf_SetNull2String(m_Rs("T13_CLUB_1_FLG")) = "1" Then 
							%>
								<%=gf_SetNull2Haifun(f_GetClubName(m_Rs("T13_CLUB_1")))%>
							<%
							End If
							%>
							<br>
						</td>

						</tr>
						<tr>
						<td nowrap class="<%=w_Class%>" width="100" align="left">
							<%
							'入部中なら表示する
							If gf_SetNull2String(m_Rs("T13_CLUB_2_FLG")) = "1" Then
							%>
								<%=gf_SetNull2Haifun(f_GetClubName(m_Rs("T13_CLUB_2")))%>
							<%
							End If
							%>
							<br>
						</td>
					</tr>

				<%If i =  w_iCnt And m_iRsCnt <> 1 Then
					'//ｽﾀｲﾙｼｰﾄのｸﾗｽを初期化
					w_Class = ""
				%>
				</table>
				</td>

				<td valign="top">
				<table class="hyo" border="1" >
					<!--ヘッダ-->
					<tr>
						<th nowrap class="header" width="15"  align="center">選<br>択</th>
						<th nowrap class="header" width="40"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
						<th nowrap class="header" width="120" align="center">氏名</th>
						<th nowrap class="header" width="100" align="center">現部活</th>
<!--
						<th nowrap class="header" width="100" align="center">中学部活</th>
-->
					</tr>
				<%End If%>

				<%m_Rs.MoveNext%>
			<%Loop%>

				</table>
				</td></tr>
			</table>
			<br>

			<table>
				<tr>
					<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="　登　録　"></td>
					<td ><input class="button" type="button" onclick="javascript:f_Cancel();" value="キャンセル"></td>
					<td ><input class="button" type="button" onclick="javascript:f_Back();" value=" 一 覧 へ "></td>
				</tr>
			</table>

		<%Exit Do%>
	<%Loop%>

	<!--値渡し用-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">
	<input type="hidden" name="chkMax"    value="<%=j%>">
	<input type="hidden" name="cboGakunenCd" value="<%=m_iGakunen%>">
	<input type="hidden" name="cboClassCd"   value="<%=m_iClassNo%>">
	<input type="hidden" name="txtTyuClubCd"   value="<%=m_sTyuClubCd%>">

	</form>
	</center>
	</body>
	</html>
<%
End Sub 
%>

