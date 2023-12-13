<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績参照（教官側）
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0800/default.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2003/05/13 廣田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Dim  m_iNendo   		'年度
	Dim  m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ
	Dim  m_Rs				'レコードセット
	Dim  m_RecCnt			'レコードカウント（科目のカウント）
	Dim  m_sGakuseiNo		'対象学生の学生番号
	Dim  m_IppanCnt			'一般　　　結合行数
	Dim  m_SenmonCnt		'専門　　　結合行数
	Dim  m_Ippan_H			'一般必修　結合行数
	Dim  m_Senmon_H			'専門必修　結合行数
	Dim  m_Ippan_S			'一般選択　結合行数
	Dim  m_Senmon_S			'専門選択　結合行数
	Dim  m_bKamokuKBN		'科目区分タイトル表示フラグ
	Dim  m_bHissenKBN		'必選区分タイトル表示フラグ

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

	Dim w_sWinTitle
	Dim w_sMsgTitle
	Dim w_sMsg
	Dim w_sRetURL
	Dim w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="成績参照"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'// 権限チェックに使用
'		Session("PRJ_No") = "SEI0800"

		'// 不正アクセスチェック
		Call gf_userChk(Session("PRJ_No"))

		'//ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

		'// 該当学生成績データ取得
		If Not f_GetGakResult() Then m_bErrFlg = True : Exit Do

		'// 該当者がいない場合
		If m_Rs.EOF Then
			Call gs_showWhitePage("個人履修データが存在しません。","成績参照")
			Exit Do
		End If

		'// ページを表示
		Call showPage()
	    Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
    
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

	m_iNendo     = Session("NENDO")				'処理年度
	m_sGakuseiNo = Request("hidGakuseiNo")		'学生NO

End Sub

Function f_GetGakResult()
'********************************************************************************
'*  [機能]  ログイン教官が担当するクラスの学生一覧を取得する
'*  [引数]  なし
'*  [戻値]  True / False
'*  [説明]  
'********************************************************************************
	On Error Resume Next
	Err.Clear

	Dim w_sSQL
	Dim w_sKamokuCD
	Dim w_lRowCnt
	Dim w_lRowCnt2

	f_GetGakResult = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T16_KAMOKU_KBN       AS KAMOKU_KBN,"
	w_sSQL = w_sSQL & " 	T16_HISSEN_KBN       AS HISSEN_KBN,"
	w_sSQL = w_sSQL & " 	T16_COURSE_CD        AS COURSE_CD, "
	w_sSQL = w_sSQL & " 	T16_SEQ_NO           AS SEQ_NO,    "
	w_sSQL = w_sSQL & " 	T16_HAITOGAKUNEN     AS HAITOGAKUNEN,"
	w_sSQL = w_sSQL & " 	T16_KAMOKU_CD        AS KAMOKU_CD,   "
	w_sSQL = w_sSQL & " 	T16_KAMOKUMEI        AS KAMOKUMEI,   "
	w_sSQL = w_sSQL & " 	T16_HAITOTANI        AS HAITOTANI,   "
	w_sSQL = w_sSQL & " 	T16_KYOKAKEIRETU_KBN AS KYOKAKEIRETU_KBN,"
	w_sSQL = w_sSQL & " 	T16_KYOKAKEIRETU_MEI AS KYOKAKEIRETU_MEI,"
	w_sSQL = w_sSQL & " 	T16_SEI_KIMATU_K     AS SEI_KIMATU_K,    "
	w_sSQL = w_sSQL & " 	T16_HTEN_KIMATU_K    AS HTEN_KIMATU_K,   "
	w_sSQL = w_sSQL & " 	T16_HYOKA_KIMATU_K   AS HYOKA_KIMATU_K,  "
	w_sSQL = w_sSQL & " 	T16_HYOTEI_KIMATU_K  AS HYOTEI_KIMATU_K, "
	w_sSQL = w_sSQL & " 	T16_TANI_SUMI        AS TANI_SUMI,       "
	w_sSQL = w_sSQL & " 	T16_HYOKA_FUKA_KBN   AS HYOKA_FUKA_KBN,  "
	w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG       AS OKIKAE_FLG,      "
	w_sSQL = w_sSQL & " 	T16_SELECT_FLG       AS SELECT_FLG       "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 		T16_NENDO       =  " & m_iNendo
	w_sSQL = w_sSQL & " 	AND T16_GAKUSEI_NO  = '" & m_sGakuseiNo & "' "
	w_sSQL = w_sSQL & " 	AND (T16_HISSEN_KBN =  " & C_HISSEN_HIS & " OR (T16_HISSEN_KBN = " & C_HISSEN_SEN & " AND T16_SELECT_FLG= " & C_SENTAKU_YES & "))"
	w_sSQL = w_sSQL & " 	AND T16_OKIKAE_FLG  <> " & C_TIKAN_KAMOKU_MOTO	'置換元以外
	w_sSQL = w_sSQL & " UNION ALL "
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T17_KAMOKU_KBN        AS KAMOKU_KBN,"
	w_sSQL = w_sSQL & " 	T17_HISSEN_KBN        AS HISSEN_KBN,"
	w_sSQL = w_sSQL & " 	T17_COURSE_CD         AS COURSE_CD, "
	w_sSQL = w_sSQL & " 	T17_SEQ_NO            AS SEQ_NO,    "
	w_sSQL = w_sSQL & " 	T17_HAITOGAKUNEN      AS HAITOGAKUNEN,"
	w_sSQL = w_sSQL & " 	T17_KAMOKU_CD         AS KAMOKU_CD,   "
	w_sSQL = w_sSQL & " 	T17_KAMOKUMEI         AS KAMOKUMEI,   "
	w_sSQL = w_sSQL & " 	T17_HAITOTANI         AS HAITOTANI,   "
	w_sSQL = w_sSQL & " 	T17_KYOKAKEIRETU_KBN  AS KYOKAKEIRETU_KBN,"
	w_sSQL = w_sSQL & " 	T17_KYOKAKEIRETU_MEI  AS KYOKAKEIRETU_MEI,"
	w_sSQL = w_sSQL & " 	T17_SEI_KIMATU_K      AS SEI_KIMATU_K,    "
	w_sSQL = w_sSQL & " 	T17_HTEN_KIMATU_K     AS HTEN_KIMATU_K,   "
	w_sSQL = w_sSQL & " 	T17_HYOKA_KIMATU_K    AS HYOKA_KIMATU_K,  "
	w_sSQL = w_sSQL & " 	T17_HYOTEI_KIMATU_K   AS HYOTEI_KIMATU_K, "
	w_sSQL = w_sSQL & " 	T17_TANI_SUMI         AS TANI_SUMI,       "
	w_sSQL = w_sSQL & " 	T17_HYOKA_FUKA_KBN    AS HYOKA_FUKA_KBN,  "
	w_sSQL = w_sSQL & " 	T17_OKIKAE_FLG        AS OKIKAE_FLG,      "
	w_sSQL = w_sSQL & " 	T17_SELECT_FLG        AS SELECT_FLG       "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T17_RISYUKAKO_KOJIN, "
	w_sSQL = w_sSQL & " 	(SELECT "
	w_sSQL = w_sSQL & " 		T13_NENDO      AS NENDO, "
	w_sSQL = w_sSQL & " 		T13_GAKUSEI_NO AS GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	 FROM "
	w_sSQL = w_sSQL & " 		T13_GAKU_NEN "
	w_sSQL = w_sSQL & " 	 WHERE "
	w_sSQL = w_sSQL & " 	 		 T13_GAKUSEI_NO = '" & m_sGakuseiNo & "'"
	w_sSQL = w_sSQL & "   		AND (T13_RYUNEN_FLG =  " & C_RYUNEN_OFF & " OR T13_RYUNEN_FLG IS NULL ) "
	w_sSQL = w_sSQL & "   	) T13 "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 		 T17_NENDO      = T13.NENDO "
	w_sSQL = w_sSQL & " 	AND  T17_GAKUSEI_NO = T13.GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND (T17_HISSEN_KBN =  " & C_HISSEN_HIS & " OR (T17_HISSEN_KBN = " & C_HISSEN_SEN & " AND T17_SELECT_FLG= " & C_SENTAKU_YES & " )) "
	w_sSQL = w_sSQL & " 	AND  T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
	w_sSQL = w_sSQL & " ORDER BY "
	w_sSQL = w_sSQL & " 	KAMOKU_KBN, "
	w_sSQL = w_sSQL & " 	HISSEN_KBN, "
	w_sSQL = w_sSQL & " 	COURSE_CD,  "
	w_sSQL = w_sSQL & " 	KYOKAKEIRETU_KBN, "
	w_sSQL = w_sSQL & " 	SEQ_NO,      "
	w_sSQL = w_sSQL & " 	OKIKAE_FLG,  "
	w_sSQL = w_sSQL & " 	HAITOGAKUNEN "

	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit Function

	'変数初期化
	m_IppanCnt  = 0
	m_SenmonCnt = 0
	m_Ippan_H   = 0
	m_Senmon_H  = 0
	m_Ippan_S   = 0
	m_Senmon_S  = 0
	m_RecCnt    = 0
	w_sKamokuCD = ""

	'1件もレコードが存在しない場合
	If m_Rs.EOF then
		f_GetGakResult = True
		Exit Function
	End If

	'科目区分別にカウントを取得
	Do While Not m_Rs.EOF
		If w_sKamokuCD <> m_Rs("KAMOKU_CD") Then					'科目別
			m_RecCnt = m_RecCnt + 1									'科目数
			w_sKamokuCD = m_Rs("KAMOKU_CD")							'科目コード一時格納
			If Cint(m_Rs("KAMOKU_KBN")) = C_KAMOKU_IPPAN Then
				m_IppanCnt = m_IppanCnt + 1							'一般科目総件数
				If Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_HIS Then
					m_Ippan_H = m_Ippan_H + 1						'一般科目必修
				ElseIf Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_SEN Then
					m_Ippan_S = m_Ippan_S + 1						'一般科目選択
				End If
			ElseIf Cint(m_Rs("KAMOKU_KBN")) = C_KAMOKU_SENMON Then
				m_SenmonCnt = m_SenmonCnt + 1						'専門科目総件数
				If Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_HIS Then
					m_Senmon_H = m_Senmon_H + 1						'専門科目必修
				ElseIf Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_SEN Then
					m_Senmon_S = m_Senmon_S + 1						'専門科目選択
				End If
			End If
		End If
		m_Rs.MoveNext
	Loop

	'//ﾚｺｰﾄﾞカウント取得
'	m_RecCnt  = gf_GetRsCount(m_Rs)
	w_lRowCnt  = Cint(m_IppanCnt) + Cint(m_SenmonCnt)
	w_lRowCnt2 = Cint(m_Ippan_H) + Cint(m_Ippan_S) + Cint(m_Senmon_H) + Cint(m_Senmon_S)

	'表示する件数とタイトル結合行数が同じかどうかを判定（科目区分）※念のため
	If Cint(m_RecCnt) = Cint(w_lRowCnt) Then
		m_bKamokuKBN = True
	Else
		m_bKamokuKBN = False
	End If

	'表示する件数とタイトル結合行数が同じかどうかを判定（必選区分）※念のため
	If Cint(m_RecCnt) = Cint(w_lRowCnt2) Then
		m_bHissenKBN = True
	Else
		m_bHissenKBN = False
	End If

	m_Rs.MoveFirst

	f_GetGakResult = True

End Function

Function f_GetHissen(p_sHissen)
'********************************************************************************
'*  [機能]  必選区分名を取得
'*  [引数]  p_sHissen - 必選区分
'*  [戻値]  True / False
'*  [説明]  
'********************************************************************************
	On Error Resume Next
	Err.Clear

	Dim w_sSQL
	Dim w_Rs

	f_GetHissen = ""

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_NENDO        = " & m_iNendo & " AND "
	w_sSQL = w_sSQL & " 	M01_DAIBUNRUI_CD = " & C_HISSEN & " AND "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = " & p_sHissen

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit Function

	f_GetHissen = w_Rs("M01_SYOBUNRUIMEI")

	w_Rs.Close
	Set w_Rs = Nothing

End Function

'****************************************************************************************
'///////////////////            パーセント形式の生成関数              ///////////////////
'----------------------------------------------------------------------------------------
'[引数]:
'		pValue - 変換対象
'		pUNum  - 小数点以下表示桁数
'  [戻り値]：
'		変換後の文字列
'[備考]:
'[作成]:2003/04/10 shin
'****************************************************************************************
Function f_FormatPercent(pValue,pUNum)
	Dim wRet , wValue , wUNum
	
	on error resume next
	
	f_FormatPercent = ""

	wValue = trim(pValue)
	wUNum = trim(pUNum)
'	If gf_IsNull(wValue) Then wValue = 0
	If gf_IsNull(wValue) Then exit function
	If gf_IsNull(wUNum) Then wUNum = 0
	
	If Err.number <> 0 Then Exit Function
	wRet = FormatNumber(wValue,wUNum,-1,,-1)
	
	f_FormatPercent = wRet
	
End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	On Error Resume Next
	Err.Clear

	'// 変数定義
	Dim w_sCell						'セルのクラス設定
	Dim w_lTani						'科目別修得単位数
	Dim w_lTotalTani				'合計修得単位数
	Dim w_sKamokuCD					'科目コード
	Dim w_sKamokuNM					'科目名
	Dim w_bLoadFLG					'初期ロードフラグ
	Dim w_sKamokuKBN				'科目区分コード
	Dim w_bDispFLG					'科目区分タイトル表示済・未表示フラグ
	Dim w_bDispFLG2					'必選区分タイトル表示済・未表示フラグ（一般）
	Dim w_bDispFLG3					'必選区分タイトル表示済・未表示フラグ（専門）
	Dim w_lGakTani(5)				'学年別単位数合計
	Dim w_sSei(5)					'成績
	Dim w_sTdColor(5)				'修得・未修得別セルカラー
	Dim i
	Dim w_iGakunen

	'// 初期化
	w_sCell      = "CELL1"			'セルのクラス設定（初期値）
	w_lTani      = 0				'科目別修得単位数
	w_lTotalTani = 0				'合計修得単位数
	w_bLoadFLG   = True				'初期ロードフラグ
	w_bDispFLG   = False			'科目区分タイトル表示済・未表示フラグ
	w_bDispFLG2  = False			'必選区分タイトル表示済・未表示フラグ（一般）
	w_bDispFLG3  = False			'必選区分タイトル表示済・未表示フラグ（専門）
	w_bDispFLG4  = True  			'必選区分タイトル表示済・未表示フラグ（専門）
	w_sKamokuCD  = ""

	For i = 1 to 5
		w_sSei(i) = ""
		w_sTdColor(i) = ""
	Next

%>
<html>

<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--

	//************************************************************
	//  [機能]  表示ボタン押下
    //************************************************************
	function jf_Submit(p_i){
		with(document.frm){
			var w_Obj = eval("hidGakNo" + p_i);
			hidGakuseiNo.value = w_Obj.value;
			target = "<%=C_MAIN_FRAME%>";
			action = "sei0800_resultdef.asp";
			submit();
		}
	}

	function jf_Back(){
		with(document.frm){
			target = "<%=C_MAIN_FRAME%>";
			action = "default.asp";
			submit();
		}
	}
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body LANGUAGE="javascript">
	<center>
	<form name="frm" METHOD="post">

	<!-- TABLEリスト部 -->
	<table border="1" class="hyo" width="630">
<%
	Do While Not m_Rs.EOF

		For i = 1 to 6

			If m_Rs.EOF Then
				Exit For
			End If

			'科目変化時にHTML表示ループを抜けを行う
			If w_sKamokuCD <> m_Rs("KAMOKU_CD") Then
				w_sKamokuCD   = m_Rs("KAMOKU_CD")											'科目コードを保持
				If w_bLoadFLG = False then													'初期ロードのみ抜ける
					Exit For
				End If
				'学年別にデータを格納
				w_bLoadFLG   = False														'初期ロードフラグ
			End If

			'// 学年別にデータを格納
			w_iGakunen = Cint(m_Rs("HAITOGAKUNEN"))
			w_sSei(w_iGakunen)     = m_Rs("HYOKA_KIMATU_K")									'成績
			w_sTdColor(w_iGakunen) = "style='background : #33CCFF;'"						'履修カラー
			If Cint(m_Rs("TANI_SUMI")) = 0 Then
				w_sTdColor(w_iGakunen) = "style='background : #FF9900;'"					'未修得カラー
			End If
			w_lGakTani(w_iGakunen) = Cint(w_lGakTani(w_iGakunen)) + Cint(m_Rs("TANI_SUMI"))	'修得単位

			w_lTani      = w_lTani + Cint(gf_SetNull2Zero(m_Rs("TANI_SUMI")))				'修得単位
			w_lTotalTani = w_lTotalTani + Cint(gf_SetNull2Zero(m_Rs("TANI_SUMI")))
			w_sKamokuKBN = m_Rs("KAMOKU_KBN")												'科目区分（一般 or 専門）
			w_sHissenKBN = m_Rs("HISSEN_KBN")												'必選区分名
			w_sKamokuNM  = m_Rs("KAMOKUMEI")												'科目名

			m_Rs.MoveNext
		Next

		w_sCell = gf_IIF(w_sCell="CELL1","CELL2","CELL1")									'セルの背景色を設定
%>
		<tr>
<%
'response.write "w_sKamokuKBN = " & w_sKamokuKBN & "<br>"
'response.write "w_sHissenKBN = " & w_sHissenKBN & "<br>"
'response.write "w_bDispFLG3 = " & w_bDispFLG3 & "<br>"
		'科目区分タイトル表示時
		If m_bKamokuKBN Then
			'初期時 AND 科目区分 = 一般
			If Cint(w_sKamokuKBN) = C_KAMOKU_IPPAN Then
				'科目区分タイトル（一般科目）
				If Not w_bDispFLG Then
					w_bDispFLG = True
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_IppanCnt & " style='writing-mode:tb-rl;' nowrap>一般科目</td>"
				End If
				'必選区分タイトル（一般必修）
				If Not w_bDispFLG2 AND Cint(w_sHissenKBN) = C_HISSEN_HIS Then
					w_bDispFLG2 = True
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Ippan_H & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				'必選区分タイトル（一般選択）
				ElseIf Not w_bDispFLG3 AND Cint(w_sHissenKBN) = C_HISSEN_SEN Then
					w_bDispFLG3 = True
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Ippan_S & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				End If
			'初期時以外 AND 科目区分 = 専門
			ElseIf Cint(w_sKamokuKBN) = C_KAMOKU_SENMON Then
				'科目区分タイトル（専門科目）
				If w_bDispFLG Then
					w_bDispFLG = False
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_SenmonCnt & " style='writing-mode:tb-rl;' nowrap>専門科目</td>"
				End If
				'必選区分タイトル（専門必修）
				If w_bDispFLG2 AND Cint(w_sHissenKBN) = C_HISSEN_HIS Then
					w_bDispFLG2 = False
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Senmon_H & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				'必選区分タイトル（専門選択）
				ElseIf w_bDispFLG4 AND Cint(w_sHissenKBN) = C_HISSEN_SEN Then

'response.write "m_RecCnt = " & Cint(m_RecCnt) & "<br>"
'response.write "m_IppanCnt = " & Cint(m_IppanCnt) & "<br>"
'response.write "m_SenmonCnt = " & Cint(m_SenmonCnt) & "<br>"
'response.write "m_Ippan_H = " & Cint(m_Ippan_H) & "<br>"
'response.write "m_Ippan_S = " & Cint(m_Ippan_S) & "<br>"
'response.write "m_Senmon_H = " & Cint(m_Senmon_H) & "<br>"
'response.end


					w_bDispFLG4 = False
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Senmon_S & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				End If
			End If
		'科目区分タイトル未表示時
		Else
			'初期時
			If Not w_sDispFLG Then
				w_sDispFLG = True
				Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_RecCnt & ">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_RecCnt & ">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			End If
		End If
%>
			<td width="250" class=<%=w_sCell%>   align="left"   height="20" nowrap>　　　<%=w_sKamokuNM%></td>
			<td width="70"  class="<%=w_sCell%>" align="center" height="20" nowrap><%=f_FormatPercent(w_lTani,1)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(1)%> nowrap><%=w_sSei(1)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(2)%> nowrap><%=w_sSei(2)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(3)%> nowrap><%=w_sSei(3)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(4)%> nowrap><%=w_sSei(4)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(5)%> nowrap><%=w_sSei(5)%></td>
		</tr>
<%
		'// 初期化
		w_lTani = 0
		For i = 1 to 5
			w_sSei(i) = ""
			w_sTdColor(i) = ""
		Next

	Loop
%>
		<tr>
			<th class="header3" align="center" colspan="3" height="20">合　　計</th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lTotalTani,1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(1),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(2),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(3),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(4),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(5),1)%></th>
		</tr>
	</table>

	<p aling="center"><input type="button" class="button" value="戻　る" onclick="jf_Back();"></p>

	</center>

	<input type="hidden" name="hidGakuseiNo">

	</form>
</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
