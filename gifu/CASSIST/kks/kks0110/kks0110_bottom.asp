<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0110/kks0110_main.asp
' 機	  能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引	  数: NENDO 		 '//処理年
'			  KYOKAN_CD 	 '//教官CD
'			  GAKUNEN		 '//学年
'			  CLASSNO		 '//ｸﾗｽNo
'			  TUKI			 '//月
' 変	  数:
' 引	  渡: NENDO 		 '//処理年
'			  KYOKAN_CD 	 '//教官CD
'			  GAKUNEN		 '//学年
'			  CLASSNO		 '//ｸﾗｽNo
'			  TUKI			 '//月
' 説	  明:
'			■初期表示
'				検索条件にかなう行事出欠入力を表示
'			■登録ボタンクリック時
'				入力情報を登録する
'-------------------------------------------------------------------------
' 作	  成: 2001/07/02 伊藤公子
' 変	  更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
	Const C_SYOBUNRUICD_IPPAN = 4	'//欠席区分(0:出席,1:欠席,2:遅刻,3:早退,4:忌引,…)
	Const C_IDO_MAX_CNT = 8			'//最大移動回数(T13移動情報登録用フィールド数)
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public	m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ
	Public	m_bDaigae			 '代替留学生取得ﾌﾗｸﾞ

	'取得したデータを持つ変数
	Public m_iSyoriNen		'//処理年度
	Public m_iKyokanCd		'//教官CD
	Public m_sGakunen		'//学年
	Public m_sClassNo		'//ｸﾗｽNO
	Public m_sTuki			'//月
	Public m_sZenki_Start	'//前期開始日
	Public m_sKouki_Start	'//後期開始日
	Public m_sKouki_End 	'//後期終了日
	Public m_sEndDay		'//入力できなくなる日

	Public m_sGakki 		'//学期
	Public m_sGakki_Kbn 	'//学期区分
	Public m_sKamokuCd		'//課目CD
	Public m_sSyubetu		'//授業種別(TUJO:通常授業,TOKU:特別活動,KBTU:個別授業)
	Public m_sHissenKbn 	'//必選区分
	Public m_iTani			'//１時限の単位数
	
	'ﾚｺｰﾄﾞセット
	Public m_Rs_M			'//recordset明細情報
	Public m_Rs_D			'//recordset代替留学生
	Public m_Rs_G			'//recordset行事出欠情報

	Public m_AryHead()		'//ヘッダ情報格納配列
	Public m_iRsCnt 		'//ヘッダﾚｺｰﾄﾞ数
	Public m_iRuiKeiCnt 	'//累計カウント
	Public m_AryRuiKei()	'//累計格納配列

	Public m_iTukiKeiCnt	'//月計カウント
	Public m_AryTukiKei()	'//月計格納配列

	Public m_AryKesseki
	Public m_iSyubetu
	Public m_iSikenKbn

	Public m_sLevelFlg
	
	Public m_iShikenInsertType	'//試験実績登録期間
								'C_SIKEN_ZEN_TYU = 1 '前期中間試験
								'C_SIKEN_ZEN_KIM = 2 '前期期末試験
								'C_SIKEN_KOU_TYU = 3 '後期中間試験
								'C_SIKEN_KOU_KIM = 4 '後期期末試験
		
	
'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*	[機能]	本ASPのﾒｲﾝﾙｰﾁﾝ
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	Dim w_iRet				'// 戻り値
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="授業出欠入力"
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
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//変数初期化
		Call s_ClearParam()
		
		'// ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()
		
		'// ヘッダリスト情報取得
		w_iRet = f_Get_HeadData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'// 生徒リスト情報取得
		w_iRet = f_Get_DetailData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//生徒情報がない場合
		If m_Rs_M.EOF Then
			'//空白ページ表示
			Call showWhitePage("生徒情報がありません")
			Exit Do
		End If
		
		'//出欠明細情報取得
		w_iRet = f_Get_AbsInfo()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//出欠学期累計取得
		w_iRet = f_Get_AbsInfo_RuiKei()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'// 管理マスタより、出欠欠課の取り方を取得
		w_iRet = gf_GetKanriInfo(m_iSyoriNen,m_iSyubetu)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
		'//出欠月計取得
		w_iRet = f_Get_AbsInfo_TukiKei()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//授業日程がない場合
		If m_iRsCnt < 0 Then
			'//空白ページ表示
			Call showWhitePage("授業日程がありません")
		   Exit Do
		End If
		
		'// データ表示ページを表示
		Call showPage()

		Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gf_closeObject(m_Rs_M)
	Call gf_closeObject(m_Rs_D)
	Call gf_closeObject(m_Rs_G)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	変数初期化
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_ClearParam()
	m_iSyoriNen = ""
	m_iKyokanCd = ""
	m_sGakunen	= ""
	m_sClassNo	= ""
	m_sTuki 	= ""
	m_sGakki	= ""
	m_sKamokuCd = ""
	m_sSyubetu	= ""
	m_iTani		= 0
	m_iShikenInsertType = 0
	
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()

	m_sZenki_Start = trim(Request("Tuki_Zenki_Start"))
	m_sKouki_Start = trim(Request("Tuki_Kouki_Start"))
	m_sKouki_End   = trim(Request("Tuki_Kouki_End"))
	m_iTani = Session("JIKAN_TANI") '１時限の単位数
	
	m_iSyoriNen = trim(Request("NENDO"))
	m_iKyokanCd = trim(Request("KYOKAN_CD"))

	m_sTuki 	= trim(Request("TUKI"))
	m_sGakki	= trim(Request("GAKKI"))

	m_sSyubetu	= trim(Request("SYUBETU"))
	m_sGakunen	= trim(Request("GAKUNEN"))
	m_sClassNo	= trim(Request("CLASSNO"))
	m_sKamokuCd = trim(Request("KAMOKU_CD"))

	If m_sGakki = "ZENKI" Then
		m_sGakki_Kbn = cstr(C_GAKKI_ZENKI)
	Else
		m_sGakki_Kbn = cstr(C_GAKKI_KOUKI)
	End If

	call gf_Get_SyuketuEnd(cint(m_sGakunen),m_sEndDay)

End Sub

'********************************************************************************
'*	[機能]	日付・曜日・時間のヘッダ情報取得処理を行う
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_HeadData()

	Dim w_sSQL
	Dim w_Rs

	On Error Resume Next
	Err.Clear
	
	f_Get_HeadData = 1

	Do 

		'//日付の範囲をセット
		Call f_GetTukiRange(w_sSDate,w_sEDate)
		
		'// 授業日付、時間データ

		'// 授業種別が個人授業（KBTU）の時は代替時間割から取得する。
		'// 2001/12/18 add
		If m_sSyubetu <> "KBTU" Then 

			'// 通常、特別授業の場合
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " SELECT"
			w_sSQL = w_sSQL & vbCrLf & "  A.T32_HIDUKE,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T20_JIGEN AS JIGEN,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T20_YOUBI_CD AS YOUBI_CD"
			w_sSQL = w_sSQL & vbCrLf & " FROM"
			w_sSQL = w_sSQL & vbCrLf & " T32_GYOJI_M A"
			w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI B"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & " B.T20_YOUBI_CD = A.T32_YOUBI_CD "
			'//特別活動の場合は、すぐ次の授業の行事情報を見て、行事かどうかを判断する
			If m_sSyubetu = "TOKU"Then
				w_sSQL = w_sSQL & vbCrLf & " AND TRUNC(B.T20_JIGEN+0.5) = A.T32_JIGEN"
			Else
				w_sSQL = w_sSQL & vbCrLf & " AND B.T20_JIGEN = A.T32_JIGEN"
			End If
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO = A.T32_NENDO"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')>='" & w_sSDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')<'"  & w_sEDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO="		& cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKKI_KBN='" & m_sGakki_Kbn & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKUNEN= "	& cInt(m_sGakunen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_CLASS= " 	& cInt(m_sClassNo)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KAMOKU='"	& trim(m_sKamokuCd) & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KYOKAN='"	& m_iKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T20_YOUBI_CD,B.T20_JIGEN "
			w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T20_JIGEN"

		Else
			'// 通常、特別授業の場合
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " SELECT"
			w_sSQL = w_sSQL & vbCrLf & "  A.T32_HIDUKE,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T23_JIGEN AS JIGEN,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T23_YOUBI_CD AS YOUBI_CD"
			w_sSQL = w_sSQL & vbCrLf & " FROM"
			w_sSQL = w_sSQL & vbCrLf & " T32_GYOJI_M A"
			w_sSQL = w_sSQL & vbCrLf & " ,T23_DAIGAE_JIKAN B"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & " B.T23_YOUBI_CD = A.T32_YOUBI_CD "
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO = A.T32_NENDO"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')>='" & w_sSDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')<'"  & w_sEDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO="		& cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_GAKKI_KBN=" & m_sGakki_Kbn & " "
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KAMOKU='"	& trim(m_sKamokuCd) & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KYOKAN='"	& m_iKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T23_YOUBI_CD,B.T23_JIGEN "
			w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T23_JIGEN"
		End If
		
		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_HeadData = 99
			Exit Do
		End If

		m_iRsCnt = 0

		'=======================
		'//時間割を配列にセット
		'=======================
		If w_Rs.EOF = false Then

			i = 0
			w_sHi = ""
			w_Rs.MoveFirst
			Do Until w_Rs.EOF


				'//取得した日付の時限が休日または、行事の場合(w_bGyoji=True)ははじく
				iRet = f_Get_DateInfo(w_Rs("T32_HIDUKE"),cint(w_Rs("JIGEN")),w_bGyoji)
				If iRet <> 0 Then
					msMsg = Err.description
					f_Get_HeadData = 99
					Exit Do
				End If

				'//休日・行事以外のみデータをセット
				If w_bGyoji <> True Then

					'//配列を設定
					ReDim Preserve m_AryHead(4,i)

					'//データ格納
					If w_sHi = gf_SetNull2String(w_Rs("T32_HIDUKE")) Then
						m_AryHead(0,i) = "" 	'//月
						m_AryHead(1,i) = "" 	'//日
						m_AryHead(2,i) = "" 	'//曜日CD
					Else
						m_AryHead(0,i) = month(gf_SetNull2String(w_Rs("T32_HIDUKE")))	  '//月
						m_AryHead(1,i) = day(gf_SetNull2String(w_Rs("T32_HIDUKE"))) 	  '//日
						m_AryHead(2,i) = gf_SetNull2String(w_Rs("YOUBI_CD"))		  '//曜日CD
					End If

					m_AryHead(3,i) = replace(gf_SetNull2String(w_Rs("JIGEN")),".","$")	'//時限
					m_AryHead(4,i) = gf_SetNull2String(w_Rs("T32_HIDUKE"))					'//日付

					w_sHi = gf_SetNull2String(w_Rs("T32_HIDUKE"))
					i = i + 1

				End If

				w_Rs.MoveNext
			Loop

		End If

		'//取得したデータ数をセット
		m_iRsCnt = i-1

		'//正常終了
		f_Get_HeadData = 0
		Exit Do
	Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
   Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	取得した日付・時限が、休日または行事でないか
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_DateInfo(p_Hiduke,p_Jigen,p_bGyoji)

	Dim w_sSQL
	Dim w_Rs
	Dim w_bGyoujiFlg

	On Error Resume Next
	Err.Clear
	
	f_Get_DateInfo = 1
	w_bGyojiFlg = False

	Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT"
		w_sSQL = w_sSQL & vbCrLf & " A.T32_GYOJI_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M A"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  A.T32_NENDO=2001 "
		w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_GAKUNEN IN (" & cInt(m_sGakunen) & "," & C_GAKUNEN_ALL & ")"
		w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_CLASS IN ("   & cInt(m_sClassNo) & "," & C_CLASS_ALL   & ")"
		w_sSQL = w_sSQL & vbCrLf & "  AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')='" & p_Hiduke & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_JIGEN=" & p_Jigen
		w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_COUNT_KBN<>" & C_COUNT_KBN_JUGYO
		w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_KYUJITU_FLG<>'" & C_HEIJITU & "'"
		
		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_DateInfo = 99
			Exit Do
		End If

		If w_Rs.EOF = False Then
			'//ﾚｺｰﾄﾞがある場合は休日か、行事の日
			w_bGyojiFlg = True
		End If

		f_Get_DateInfo = 0
		Exit Do
	Loop

		'//戻り値をセット
		p_bGyoji = w_bGyojiFlg

		'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	   Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	明細情報を取得する
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_DetailData()

	Dim w_iRet

	On Error Resume Next
	Err.Clear
	
	f_Get_DetailData = 1

	Do 

		'//授業種別により処理を分岐(TUJO:通常授業,TOKU:特別活動,KBTU:個別授業)
		Select Case trim(m_sSyubetu)
		  Case "TUJO" ':通常授業

			'//通常授業取得時
			w_iRet = f_Get_Data_TUJO()
			If w_iRet <> 0 then
				Exit Do
			End If

		  Case "TOKU" ':特別活動

			'//特別授業取得時(担任ｸﾗｽ一覧)
			w_iRet = f_Get_Data_TOKU()
			If w_iRet <> 0 then
				Exit Do
			End If

		  Case "KBTU" ':個別授業

			'//個別授業取得時(課目受持ち留学生一覧)
			w_iRet = f_Get_Data_KOBETU()
			If w_iRet <> 0 then
				Exit Do
			End If

		  Case Else
			'//システムエラー
			m_sErrMsg = "パラメータが不足しています。(システムエラー)"
		End Select

		f_Get_DetailData = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	通常授業選択時クラス一覧を取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_Data_TUJO()

	Dim w_sSQL
	Dim w_Rs
	Dim w_iRet
	Dim w_sLevelFlg

	On Error Resume Next
	Err.Clear
	
	f_Get_Data_TUJO = 1
	w_sLevelFlg = ""

	Do 
		'//入学年度(=処理年度-学年+1)
		w_NyuNen = cInt(m_iSyoriNen) - cInt(m_sGakunen) + 1

		'================
		'//授業情報取得
		'================
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_HISSEN_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_LEVEL_FLG"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS,T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD = T15_RISYU.T15_GAKKA_CD AND"
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO=" 	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN="	 & cInt(m_sGakunen)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO="	 & cInt(m_sClassNo)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO="	 & w_NyuNen 		 & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_CD='" & trim(m_sKamokuCd) & "'"
		
		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_Data_TUJO = 99
			Exit Do
		End If
		
		If w_Rs.EOF = False Then
			'//レベル課目ﾌﾗｸﾞを取得
			w_sLevelFlg = w_Rs("T15_LEVEL_FLG")
			m_sLevelFlg = w_Rs("T15_LEVEL_FLG")
			'//必選区分を取得
			m_sHissenKbn =w_Rs("T15_HISSEN_KBN")
		End If
		
		'//通常課目生徒一覧取得
		w_iRet = f_Get_TUJO_Tujyo()
		If w_iRet <> 0 Then
			Exit Do
		End If
		
		'//通常授業選択時、代替留学生一覧取得
		w_iRet = f_Get_TUJO_DaigeRyugak()
		If w_iRet <> 0 Then
			Exit Do
		End If
		
		f_Get_Data_TUJO = 0
		Exit Do
	Loop
	
	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	レベル別課目生徒一覧取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_TUJO_LevelBetu()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_TUJO_LevelBetu = 1

	Do 

		'// レベル別課目選択生徒一覧取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI  AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_SELECT_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_OKIKAE_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI ,T16_RISYU_KOJIN,T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_GAKUSEI_NO = T16_RISYU_KOJIN.T16_GAKUSEI_NO AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO="			& cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_NENDO="		   & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_HAITOGAKUNEN="   & cInt(m_sGakunen)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD='"	   & m_sKamokuCd	   & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_LEVEL_KYOUKAN='" & m_iKyokanCd	   & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI"

		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_TUJO_LevelBetu = 99
			Exit Do
		End If

		f_Get_TUJO_LevelBetu = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	通常課目生徒一覧取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_TUJO_Tujyo()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_TUJO_Tujyo = 1

	Do 

		'// 通常課目生徒一覧
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_SELECT_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_OKIKAE_FLG,"
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_LEVEL_KYOUKAN,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN ,T16_RISYU_KOJIN,T11_GAKUSEKI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T16_RISYU_KOJIN.T16_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO = T16_RISYU_KOJIN.T16_NENDO  AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO="	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS="	 & cInt(m_sClassNo)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD='" & m_sKamokuCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI "
		
		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_TUJO_Tujyo = 99
			Exit Do
		End If
		
		f_Get_TUJO_Tujyo = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	特別授業選択時、担任クラス一覧取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_Data_TOKU()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_Data_TOKU = 1

	Do 

		'// 担任クラス一覧
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN," 
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "   T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN,T11_GAKUSEKI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS=" & cInt(m_sClassNo)
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI "
		
		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_Data_TOKU = 99
			Exit Do
		End If
		
		f_Get_Data_TOKU = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	個別授業選択時、留学生一覧取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_Data_KOBETU()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_Data_KOBETU = 1

	Do 

		'// 留学生一覧取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO  AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN ,T11_GAKUSEKI ,T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO = T23_DAIGAE_JIKAN.T23_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO = T13_GAKU_NEN.T13_GAKUSEKI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO=" & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN='" & m_sGakki_Kbn	  & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KAMOKU='" & m_sKamokuCd & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KYOKAN='" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI"
		
		'response.write "w_sSQL =" & w_sSQL & "<BR>"
		
		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_Data_KOBETU = 99
			Exit Do
		End If
		
		f_Get_Data_KOBETU = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	通常授業選択時、代替留学生一覧取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_TUJO_DaigeRyugak()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_TUJO_DaigeRyugak = 1

	Do 

		'//代替取得ﾌﾗｸﾞ
		m_bDaigae = True

		'// 留学生一覧取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KAMOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KYOKAN, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO  AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN,"
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO = T13_GAKU_NEN.T13_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUNEN = T13_GAKU_NEN.T13_GAKUNEN AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_CLASS = T13_GAKU_NEN.T13_CLASS AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO = T13_GAKU_NEN.T13_GAKUSEKI_NO  AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO="		& cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN='" & m_sGakki_Kbn		& "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KAMOKU='"	& m_sKamokuCd		& "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KYOKAN='"	& m_iKyokanCd		& "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI"
		
		iRet = gf_GetRecordset(m_Rs_D, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_TUJO_DaigeRyugak = 99
			Exit Do
		End If

		f_Get_TUJO_DaigeRyugak = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	移動ありの場合移動状況の取得
'*	[引数]	p_Gakusei_No:学績NO
'*			p_Date		:授業実施日
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_IdouInfo(p_Gakusei_No,p_Date)

	Dim w_sSQL
	Dim w_Rs
	Dim w_IdoFlg
	Dim w_sKubunName

	On Error Resume Next
	Err.Clear

	w_IdoFlg = False

	Do

		'// 明細データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
		w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cint(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO='" & p_Gakusei_No & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			Exit Do
		End If

		If w_Rs.EOF = false Then

			i = 1
			Do Until i> cint(C_IDO_MAX_CNT)    '//C_IDO_MAX_CNT = 8…最大移動回数

				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
					Exit Do
				End If

				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > p_Date  Then
					Exit Do
				End If
				i = i + 1
			Loop

			If i = 1 then
				'//最初の移動日が授業日より未来の場合、授業日に移動状態ではない
				w_sKubunName = ""
			Else

				Select Case Trim(w_Rs("T13_IDOU_KBN_" & i-1))
					Case cstr(C_IDO_FUKUGAKU),cstr(C_IDO_TEI_KAIJO)  '//C_IDO_FUKUGAKU=3:復学、C_IDO_TEI_KAIJO=5:停学解除
						w_sKubunName = ""
					Case Else
						'//移動理由を取得(区分マスタ、大分類=C_IDO)
						w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),m_iSyoriNen,w_sKubunName)
						If w_bRet<> True Then
							Exit Do
						End If
				End Select
			End If

		End If

		Exit Do
	Loop

	f_Get_IdouInfo = w_sKubunName

	Call gf_closeObject(w_Rs)

	Err.Clear

End Function

'********************************************************************************
'*	[機能]	教科別出欠データを取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_AbsInfo()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_AbsInfo = 1

	Do 

		'//月の範囲をセット
		Call f_GetTukiRange(w_sSDate,w_sEDate)

		'// 出欠データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIMU_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO = M01_KUBUN.M01_NENDO(+) AND  "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN = M01_KUBUN.M01_SYOBUNRUI_CD(+) AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO="	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE>='"  & w_sSDate & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE< '"  & w_sEDate & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KAMOKU='" & m_sKamokuCd		 & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KYOKAN='" & m_iKyokanCd		 & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_DAIBUNRUI_CD=" & C_KESSEKI	'//C_KESSEKI = 19 大分類(欠席区分)
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIMU_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
		
		iRet = gf_GetRecordset(m_Rs_G, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_AbsInfo = 99
			Exit Do
		End If

		f_Get_AbsInfo = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	個人の行事出欠データを返す
'*	[引数]	p_Date		:日付
'*			p_Gakuseki	:学績
'*			p_Jigen 	:時限
'*	[戻値]	p_sSyuketu	:出欠ｺｰﾄﾞ(データなしの場合は0(出席)を返す)
'*			p_sSyuketu_R:出欠略称
'*			p_bJim		:True=事務入力 False=授業担当教官入力
'*	[説明]	
'********************************************************************************
Function f_Get_Syuketu(p_Date,p_Gakuseki,p_Jigen,p_bJim,p_sSyuketu,p_sSyuketu_R)

	Dim w_sSyuketu

	On Error Resume Next
	Err.Clear

	p_sSyuketu = ""
	p_sSyuketu_R = "　"
	p_bJim = False
	Do
		If m_Rs_G.EOF = False Then
			m_Rs_G.MoveFirst
			Do Until m_Rs_G.EOF 
				If p_Date = m_Rs_G("T21_HIDUKE") Then
					If trim(p_Gakuseki) = trim(m_Rs_G("T21_GAKUSEKI_NO")) Then

						If cstr(replace(p_Jigen,"$",".")) = cstr(m_Rs_G("T21_JIGEN")) Then

							'//出欠ｺｰﾄﾞ
							p_sSyuketu = m_Rs_G("T21_SYUKKETU_KBN")

							If cstr(p_sSyuketu) = cstr(C_KETU_SYUSSEKI) Then
								p_sSyuketu_R = "　"
							Else
								p_sSyuketu_R = m_Rs_G("M01_SYOBUNRUIMEI_R")
							End If 

							'//事務入力されたデータは、変更不可とする
							'//入力ﾌﾗｸﾞ(0:教官 1:事務)
							If cstr(gf_SetNull2String(m_Rs_G("T21_JIMU_FLG"))) = cstr(C_JIMU_FLG_JIMU) then
								p_bJim = True
							End If

							Exit Do
						End If
					End If
				End If
				m_Rs_G.MoveNext
			Loop
			m_Rs_G.MoveFirst
		End If

		Exit Do

	Loop

	Err.Clear

End Function

'********************************************************************************
'*	[機能]	一番近い未来の試験区分を取得
'*	[引数]	なし
'*	[戻値]	p_iSikenKbn 試験区分
'*	[説明]	
'********************************************************************************
Function f_GetSikenKbn(p_iSikenKbn)
	Dim w_sSQL
	Dim w_Rs
	Dim w_iRet

	Dim w_sDate

	On Error Resume Next
	Err.Clear
	
	f_GetSikenKbn = 1
	p_iSikenKbn = ""

	Do 
		'2001/12/17 Add >
		if Cint(m_sTuki) < 4 then
			w_sDate = gf_YYYY_MM_DD((Cint(m_iSyoriNen) + 1) & "/" & m_sTuki & "/01","/")
		Else
			w_sDate = gf_YYYY_MM_DD(m_iSyoriNen & "/" & m_sTuki & "/01","/")
		end if
		if m_sKouki_Start > w_sDate then
			w_sDate = m_sKouki_Start
		End if

		'//試験管理マスタより試験区分を取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SIKEN_KBN "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & m_sGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_JISSI_SYURYO>='" & w_sDate & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY T24_SIKEN_NITTEI.T24_SIKEN_KBN"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetSikenKbn = 99
			Exit Do
		End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_iSikenKbn = w_Rs("T24_SIKEN_KBN")
		End If

		'//データが取得できないとき(後期期末後)は、0にする
		If p_iSikenKbn = "" Then
			p_iSikenKbn = 0
		End If

		f_GetSikenKbn = 0

		Exit Do
	Loop

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	出欠累計データを取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_AbsInfo_RuiKei()

	Dim w_sSQL
	Dim rs
	Dim w_GakusekiNo
	
	On Error Resume Next
	Err.Clear
	
	f_Get_AbsInfo_RuiKei = 1

	Do 
		
		'//一番近い未来の試験区分を取得する
		w_iRet = f_GetSikenKbn(m_iSikenKbn)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Get_AbsInfo_RuiKei = 99
			Exit Do
		End If
		
		'//後期期末以降の場合
		If cint(m_iSikenKbn) = 0 Then
			'w_iSiken = 4
			w_iSiken = 5
		Else
			w_iSiken = m_iSikenKbn
		End If
		
		'//最初の生徒の学籍番号を取得
		if not m_Rs_M.EOF then
			w_GakusekiNo = m_Rs_M("GAKUSEKI")
			m_Rs_M.movefirst
		end if
		
		'//前の試験と次の試験間の開始日、終了日を取得
		w_bRtn = gf_GetStartEnd("kks",m_iSyoriNen,m_sSyubetu,cint(w_iSiken),m_sGakunen,m_sClassNo,m_sKamokuCd,w_sKaisibi,w_sSyuryobi,m_iShikenInsertType)
		If w_bRtn <> True Then
			Exit Function
		End If
		
		'// 出欠データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    Count(A.T21_SYUKKETU_KBN) AS KAISU,"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "    T21_SYUKKETU A "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_NENDO="	& cInt(m_iSyoriNen) & " AND "
		
		w_sSQL = w_sSQL & vbCrLf & "    to_date(A.T21_HIDUKE,'YYYY/MM/DD') >= '" & w_sKaisibi & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "    to_date(A.T21_HIDUKE,'YYYY/MM/DD') <=  '" & w_sSyuryobi & "' AND "
		
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_KAMOKU='" & m_sKamokuCd & "' AND "
		'w_sSQL = w_sSQL & vbCrLf & "    A.T21_KYOKAN='" & m_iKyokanCd & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN IN (" & C_KETU_KEKKA & "," & C_KETU_TIKOKU & "," & C_KETU_SOTAI & "," & C_KETU_KEKKA_1 & ")"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T21_GAKUSEKI_NO"
		
		If gf_GetRecordset(rs, w_sSQL) <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_AbsInfo_RuiKei = 99
			Exit Do
		End If

		If rs.EOF = false  Then
			m_iRuiKeiCnt = gf_GetRsCount(rs) - 1
			ReDim Preserve m_AryRuiKei(2,m_iRuiKeiCnt)
			
			'//初期化
			For i=0 to m_iRuiKeiCnt
				For j=0 to 2
					m_AryRuiKei(j,i)=0
				Next
			Next


			i = 0
			Do Until rs.EOF
				If w_GakuNo <> trim(rs("T21_GAKUSEKI_NO")) Then
					If w_GakuNo <> "" Then
						i = i + 1
					End If
					w_GakuNo = trim(rs("T21_GAKUSEKI_NO"))
					m_AryRuiKei(0,i) = w_GakuNo
				End If

				Select case cstr(rs("T21_SYUKKETU_KBN"))
					case cstr(C_KETU_KEKKA) 	'//欠課数
						m_AryRuiKei(1,i) = m_AryRuiKei(1,i) + cint(rs("KAISU")) * m_iTani
					case cstr(C_KETU_TIKOKU)	'//遅刻数
						m_AryRuiKei(2,i) = m_AryRuiKei(2,i) + cint(rs("KAISU"))
					case cstr(C_KETU_SOTAI)		'//早退数
						m_AryRuiKei(2,i) = m_AryRuiKei(2,i) + cint(rs("KAISU"))
					case cstr(C_KETU_KEKKA_1) 	'//欠課数（１欠課分）
						m_AryRuiKei(1,i) = m_AryRuiKei(1,i) + cint(rs("KAISU"))
				End Select
				
				If w_GakuNo <> trim(rs("T21_GAKUSEKI_NO")) Then
					i = i + 1
					m_AryRuiKei(0,i) = trim(rs("T21_GAKUSEKI_NO"))
				End If
				
				rs.MoveNext
			Loop
		End If
		
		f_Get_AbsInfo_RuiKei = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function



'********************************************************************************
'*	[機能]	出欠月計データを取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	
'********************************************************************************
Function f_Get_AbsInfo_TukiKei()

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear
	
	f_Get_AbsInfo_TukiKei = 1

	Do 
		'//月の範囲をセット
		Call f_GetTukiRange(w_sSDate,w_sEDate)
		
		'// 出欠データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    Count(A.T21_SYUKKETU_KBN) AS CNT,"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "    T21_SYUKKETU A"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_NENDO="	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_HIDUKE >= '" & w_sSDate	 & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_HIDUKE < '"  & w_sEDate	 & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_KAMOKU='"    & m_sKamokuCd & "' AND"
		'w_sSQL = w_sSQL & vbCrLf & "    A.T21_KYOKAN='"    & m_iKyokanCd & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN IN ('" & cstr(C_KETU_KEKKA) & "','" & cstr(C_KETU_TIKOKU) & "','" & cstr(C_KETU_SOTAI) & "','" & cstr(C_KETU_KEKKA_1) & "')"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T21_GAKUSEKI_NO"
		
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			msMsg = Err.description
			f_Get_AbsInfo_TukiKei = 99
			Exit Do
		End If

		If rs.EOF= false  Then

			'//ﾚｺｰﾄﾞカウント取得
			m_iTukiKeiCnt = gf_GetRsCount(rs) - 1

			ReDim Preserve m_AryTukiKei(2,cInt(m_iTukiKeiCnt))

			'//初期化
			For j=0 to 2
				For i=0 to m_iTukiKeiCnt
					m_AryTukiKei(j,i)=0
				Next
			Next

			i = 0
			Do Until rs.EOF

				If w_GakuNo <> trim(rs("T21_GAKUSEKI_NO")) Then
					If w_GakuNo <> "" Then
						i = i + 1
					End If
					w_GakuNo = trim(rs("T21_GAKUSEKI_NO"))
					m_AryTukiKei(0,i) = w_GakuNo
				End If

				Select case cstr(rs("T21_SYUKKETU_KBN"))
					case cstr(C_KETU_KEKKA) 	'//欠課数
						m_AryTukiKei(1,i) = m_AryTukiKei(1,i) + cint(rs("CNT")) * m_iTani
					case cstr(C_KETU_TIKOKU)	'//遅刻数
						m_AryTukiKei(2,i) = m_AryTukiKei(2,i) + cint(rs("CNT"))
					case cstr(C_KETU_SOTAI)		'//早退数
						m_AryTukiKei(2,i) = m_AryTukiKei(2,i) + cint(rs("CNT"))
					case cstr(C_KETU_KEKKA_1) 	'//欠課数（１欠課分）
						m_AryTukiKei(1,i) = m_AryTukiKei(1,i) + cint(rs("CNT"))
				End Select
				
				rs.MoveNext
			Loop
			
		End If
		
		f_Get_AbsInfo_TukiKei = 0
		Exit Do
	Loop
	
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*	[機能]	出欠区分と名称を取得
'*	[引数]	なし
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	出欠入力のJAVASCRIPT作成
'********************************************************************************
Function f_Get_SYUKETU_KBN(p_MaxNo)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear
	
	f_Get_SYUKETU_KBN = 1

	Do 
		'// 明細データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_KESSEKI) & " AND "
		'//C_SYOBUNRUICD_IPPAN = 4	'//欠席区分(0:出席,1:欠席,2:遅刻,3:早退,4:忌引,…)
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD IN ("
		w_sSQL = w_sSQL & vbCrLf & " " & C_KETU_SYUSSEKI & " "
		w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_KEKKA & " "
		w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_TIKOKU & " "
		w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_SOTAI & " "
		
		If m_iTani > 1 then '１時限が１欠課より大きい場合は、１欠課の区分を出す
			w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_KEKKA_1 & " "
		End If

		w_sSQL = w_sSQL & vbCrLf & " ) "
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY M01_KUBUN.M01_SYOBUNRUI_CD"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_SYUKETU_KBN = 99
			Exit Do
		End If

		i=0
		If rs.EOF = True Then
			response.write ("var ary = new Array(0);")
			response.write ("var aryCD = new Array(0);")

			response.write ("aryCD[0] = '';")
			response.write ("ary[0] = '';")
		Else

			'//ﾚｺｰﾄﾞカウント取得
			w_iCnt = gf_GetRsCount(rs) - 1
			response.write ("var aryCD = new Array(" & w_iCnt & ");") & vbCrLf
			response.write ("var ary = new Array(" & w_iCnt & ");") & vbCrLf

			Do Until rs.EOF
'			 response.write ("var ary[" & i & "] = new Array(1);") & vbCrLf
				If i = 0 Then
					response.write ("aryCD[0] = 0;") & vbCrLf
					response.write ("ary[0] = '　';") & vbCrLf
				Else
'					 response.write ("ary[" & rs("M01_SYOBUNRUI_CD") &	"] = '" & rs("M01_SYOBUNRUIMEI_R") & "';") & vbCrLf

					'下の文に修正 2001/10/29
					'aryCD=小分類コード
					'ary=小分類略称
					response.write ("aryCD[" & i &	"] = '" & rs("M01_SYOBUNRUI_CD") & "';") & vbCrLf
					response.write ("ary[" & i &  "] = '" & rs("M01_SYOBUNRUIMEI_R") & "';") & vbCrLf
				End If

				i=i+1
				rs.MoveNext
			Loop

		End If

		p_MaxNo = w_iCnt

		f_Get_SYUKETU_KBN = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)
	Err.Clear

End Function

'********************************************************************************
'*	[機能]	配列初期化
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_AryInit(p_iRecCount)

	For j=0 to 4
		For i=0 to p_iRecCount
			m_AryHead(j,i) = ""
		Next
	Next

End Sub

'********************************************************************************
'*	[機能]	月の検索条件を作成(7月…　"MONTH>=2001/07/01 AND MONTH<2001/08/01" として使用)
'*	[引数]	なし
'*	[戻値]	p_sSDate
'*			p_sEDate
'*	[説明]	
'********************************************************************************
Function f_GetTukiRange(p_sSDate,p_sEDate)

	p_sSDate = ""
	p_sEDate = ""

	If m_sGakki = "ZENKI" Then
		w_iNen = cint(m_iSyoriNen)

		'//開始日
		If cint(month(m_sZenki_Start)) = Cint(m_sTuki) Then
			p_sSDate = m_sZenki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki),2) & "/01"
		End If

		'//終了日
		If cint(month(m_sKouki_Start)) = Cint(m_sTuki) Then
			p_sEDate = m_sKouki_Start
		Else 
			If Cint(m_sTuki) = 12 Then
				p_sEDate = cstr(w_iNen+1) & "/01/01"
			Else
				p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki+1),2) & "/01"
			End If
		End If

	Else
		'//後期の年
		If cint(m_sTuki) <=4 Then
			w_iNen = cint(m_iSyoriNen) + 1
		Else
			w_iNen = cint(m_iSyoriNen)
		End If

		'//開始日
		If cint(month(m_sKouki_Start)) = Cint(m_sTuki) Then
			p_sSDate = m_sKouki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki),2) & "/01"
		End If

		'//終了日
		If cint(month(m_sKouki_End)) = Cint(m_sTuki) Then
			p_sEDate = DateAdd("d",1,m_sKouki_End)
		Else 
			If Cint(m_sTuki) = 12 Then
				p_sEDate = cstr(w_iNen+1) & "/01/01"
			Else
				p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki+1),2) & "/01"
			End If
		End If

	End If

End Function

'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showPage()
Dim w_sIduoRiyu
Dim w_bJimInsData	'//事務入力ﾌﾗｸﾞ
Dim w_bNoSelect		'//教科非選択ﾌﾗｸﾞ
Dim w_bNoChange		'//変更不可ﾌﾗｸﾞ
Dim w_bEndFLG		'//すべて変更不可の場合TRUE

	On Error Resume Next
	Err.Clear

	w_bEndFLG = True
%>
	<html>
	<head>
	<title>行事用出欠入力</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->

	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {

		//スクロール同期制御
		parent.init();
	if(location.href.indexOf('#')==-1)
 	{
		//ヘッダ部を表示submit
		//document.frm.target = "middle";
		document.frm.target = "topFrame";
		document.frm.action = "kks0110_middle.asp"
		document.frm.submit();
	}

		return;

	}

	//************************************************************
	//	[機能]	出欠入力
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function chg(chgInp) {

		no = 0;
		<%
		'//出欠区分を取得
		Call f_Get_SYUKETU_KBN(w_MaxNo)
		%>

		str = chgInp.value;
		for(i=0; i<<%=w_MaxNo+1%>; i++){
			if (ary[i]==str){
				break;
			}
		};

		no = i + 1;
		if (no > <%=w_MaxNo%>) no = 0;
		chgInp.value = ary[no];

		//隠しフィールドにデータをセット
		var obj=eval("document.frm.hid"+chgInp.name);
		obj.value=aryCD[no];
		return;
	}
	//************************************************************
	//	[機能]	登録ボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Touroku(){

		if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
		   return ;
		}

		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("登録しています・・・・　　しばらくお待ちください")%>"

		//リスト情報をsubmit
		document.frm.target = "main";
		document.frm.action = "./kks0110_edt.asp"
		document.frm.submit();
		return;
	}

	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Cancel(){
		//初期ページを表示
		parent.document.location.href="default.asp"
	}

	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="window_onload()">
	<form name="frm" method="post" onClick="return false;">

	<center>
	<%Do %>
		<%If m_iRsCnt < 0 Then%>
			<br><br>
			<span class="msg">授業日程がありません</span>
			<%Exit Do%>
		<%End If%>

		<%If m_Rs_M.EOF Then%>
			<br><br>
			<span class="msg">履修データがありません</span>
			<%Exit Do%>
		<%End If%>

		<%'//時限を隠し項目にｾｯﾄ%>
		<%for i = 0 to m_iRsCnt%>
			<input type="hidden" name="JIKANWARI" value="<%=m_AryHead(4,i) & "_" & m_AryHead(3,i)%>">
		<%Next%>

		<table >
		<tr>
			<td align="center" valign="top">
			<table class="hyo"	border="1" >

			<%

			Dim w_sSentaku

			'//明細入力欄表示
			If m_Rs_M.EOF = False Then

				Do Until m_Rs_M.EOF

					'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
					Call gs_cellPtn(w_Class) 
					%>
					<tr>
						<td class="<%=w_Class%>" align="center" height="28" nowrap width="50"><%=m_Rs_M("GAKUSEKI")%>
							<input type="hidden" name="GAKUSEI" value=<%=m_Rs_M("GAKUSEI")%>>
						</td>
						<td class="<%=w_Class%>" align="left" height="28" nowrap width="150"><%=m_Rs_M("SIMEI")%></td>
					<%
					'//教科非選択ﾌﾗｸﾞ
					w_bNoSelect = False

					'//通常授業選択時(特別授業、個別授業以外のみ)
					If m_sSyubetu = "TUJO" then

						'//必選区分が2の科目(選択科目)のとき、選択可否を判別(選択する=1、選択しない=0)選択しない場合は、出欠入力不可
						If cstr(m_sHissenKbn) = cstr(C_HISSEN_SEN) Then

							'//選択可否を判別
							If cstr(gf_SetNull2Zero(m_Rs_M("T16_SELECT_FLG"))) = cstr(gf_SetNull2Zero(C_SENTAKU_NO)) Then
								w_bNoSelect = True
							End If
						Else
							'//置換科目ﾌﾗｸﾞが1(0:通常,1:置換元,2:置換先)の場合は、出欠入力不可
							If cstr(gf_SetNull2Zero(m_Rs_M("T16_OKIKAE_FLG"))) = cstr(gf_SetNull2Zero(C_TIKAN_KAMOKU_MOTO)) Then
								w_bNoSelect = True
							End If

	
							'ﾚﾍﾞﾙ別科目
							If cstr(trim(m_sLevelFlg)) = cstr(C_LEVEL_YES) Then

								'//ﾚﾍﾞﾙ別科目 ﾚﾍﾞﾙ別教官コードがNULLなら、出欠入力不可 ito
								If isNull(m_Rs_M("T16_LEVEL_KYOUKAN")) = True Then
									w_bNoSelect = True
								Else
									If m_Rs_M("T16_LEVEL_KYOUKAN") <> m_iKyokanCd Then
										w_bNoSelect = True
									End If
								End If
							End If
					
						End If

					End If

					For i = 0 to m_iRsCnt

						'//変更不可ﾌﾗｸﾞ初期化
						w_bNoChange = False

						'//生徒が教科を選択してない場合
						If w_bNoSelect = True Then
							w_bNoChange = True
						End If

						'//移動状況の考慮(T13_IDOU_NUMが1以上の場合は移動状況を判別する)移動中の場合は、出欠入力不可
						w_sIduoRiyu = ""
						If gf_SetNull2Zero(m_Rs_M("T13_IDOU_NUM")) > 0 Then 
							w_sIduoRiyu = f_Get_IdouInfo(m_Rs_M("GAKUSEI"),m_AryHead(4,i))
						End If

						'//移動中でない場合出欠データを取得
						If Trim(w_sIduoRiyu) = "" Then
							'//日付、学績NO(5桁),時限,事務入力(事務入力ﾌﾗｸﾞが1の場合は入力不可)より、出欠情報等を取得
							Call f_Get_Syuketu(m_AryHead(4,i),m_Rs_M("GAKUSEKI"),m_AryHead(3,i),w_bJimInsData,w_Syuketu,w_Syuketu_R)
							'//事務入力ﾌﾗｸﾞが1の場合は変更不可とする
							If w_bJimInsData = True Then
								w_bNoChange = True
							End If
						End If
						'//移動中(入力不可)
						If w_sIduoRiyu <> "" Then%>
							<td align="center" class="NOCHANGE" height="28" nowrap width="30" ><%=w_sIduoRiyu%><br>
							<input type="hidden" name="hidKBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
						<%
						'//事務入力 OR 選択していない(入力不可) ito
						Else

							If w_bNoChange = True Then%>
								<td align="center" class="NOCHANGE" height="28" nowrap	width="30" ><%=w_Syuketu_R%><br>
								<input type="hidden" name="hidKBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
							<%
							'//変更・入力可データ
							Else%>
								
								<% '変更可能期間の場合
									w_bEndFLG = False %>
									<td class="<%=w_Class%>" align="center" width="30" height="28" nowrap>
									<input type="button" class="<%=w_Class%>" name="KBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu_R%>"	style="border-style:none" style="text-align:center" tabindex="-1" onclick="return chg(this)">
									<input type="hidden" name="hidKBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu%>"></td>

							
							<%End If%>
					 <%End If%>

					<%Next%>
					</tr>
					<%m_Rs_M.MoveNext%>
				<%Loop%>

					<%m_Rs_M.MoveFirst%>
			<%End If%>

			<%

			'======================================
			'//代替留学生の追加
			If m_bDaigae = True Then
				
				If m_Rs_D.EOF = false Then

					Do Until m_Rs_D.EOF
					'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
					Call gs_cellPtn(w_Class) 

						%>
						<tr>
							<td class="<%=w_Class%>" align="center" height="28" nowrap><%=m_Rs_D("GAKUSEKI")%>
								<input type="hidden" name="GAKUSEI" value=<%=m_Rs_D("GAKUSEI")%>>
							</td>
							<td class="<%=w_Class%>" align="center" height="28" nowrap><%=m_Rs_D("SIMEI")%></td>
						<%
						for i = 0 to m_iRsCnt
							w_bNoChange = False

							'//入力許可・非許可の判定
							'//移動状況の考慮(T13_IDOU_NUMが1以上の場合は移動状況を判別する)移動中の場合は、出欠入力不可
							w_sIduoRiyu = ""
							If gf_SetNull2Zero(m_Rs_D("T13_IDOU_NUM")) > 0 Then 
								w_sIduoRiyu = f_Get_IdouInfo(m_Rs_M("GAKUSEI"),m_AryHead(4,i))
							End If

							'//日付、学績NO(5桁),時限,担任入力(担任入力ﾌﾗｸﾞが1の場合は担任のみ入力可)
							Call f_Get_Syuketu(m_AryHead(4,i),m_Rs_D("GAKUSEKI"),m_AryHead(3,i),w_bJimInsData,w_Syuketu,w_Syuketu_R)

							'//事務入力ﾌﾗｸﾞが1の場合は変更不可とする
							If w_bJimInsData = True Then
								w_bNoChange = w_bJimInsData
							End If

							'//時間割の時限と代替時間割の時限が一致しているか
							'//一致してない場合は留学生はその時限を選択していないとみなし、出欠入力を不可とする
							If w_bNoChange = False Then
								If cstr(replace(m_AryHead(3,i),"$",".")) <> cstr(m_Rs_D("T23_JIGEN")) Then
									w_bNoChange = True
								End If
							End If

							'//移動中(入力不可)
							If w_sIduoRiyu <> "" Then%>
								<td align="center" class=NOCHANGE height="28" nowrap><%=w_sIduoRiyu%><br></td>
								<input type="hidden" name="hidKBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
							<%
							'//事務入力 OR 選択していない(入力不可)
							ElseIf w_bNoChange = True Then%>
								<td align="center" class=NOCHANGE height="28" nowrap><%=w_Syuketu_R%><br></td>
								<input type="hidden" name="hidKBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
							<%Else%>
								<% '変更可能期間の場合 %>
								<% w_bEndFLG = False %>
									<td class="<%=w_Class%>" align="center"  width="30" height="28"  nowrap>
									<input type="button" class="<%=w_Class%>" name="KBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu_R%>" style="border-style:none" style="text-align:center" tabindex="-1" onclick="return chg(this)">
									<input type="hidden" name="hidKBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu%>"></td>
								<% '変更可能期間でない場合 %>


							<%End If

						Next

						m_Rs_D.MoveNext
					Loop
					m_Rs_D.MoveFirst
				End If
				w_Class=""
			End If
			'======================================
			%>

			</table>

		</td>
		<td width="10"><br></td>
		<td align="center" valign="top" width="120"  nowrap>

			<!--月・学期の欠席及び遅刻数累計-->
			<table	class="hyo" border="1" width="120">

			<%If m_Rs_M.EOF = False Then
				w_Class = ""
				Do Until m_Rs_M.EOF
					'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
					Call gs_cellPtn(w_Class) 
					%>
					<tr>
					<%
					w_sGakusekiNo = m_Rs_M("GAKUSEKI")

					'//初期化
					w_TukiTikoku = 0   '//月遅刻
					w_TukiKekka  = 0   '//月欠課
					w_RuiTikoku  = 0   '//累計遅刻
					w_RuiKekka	 = 0   '//累計欠課
					
					'//月計を取得
					
					For i=0 To cInt(m_iTukiKeiCnt)
						If w_sGakusekiNo=m_AryTukiKei(0,i) Then
							w_TukiKekka  = m_AryTukiKei(1,i)	'//欠課数
							w_TukiTikoku = m_AryTukiKei(2,i)	'//遅刻数
							Exit For
						End If
					Next

					'//累計を取得
					For i=0 To m_iRuiKeiCnt
						If w_sGakusekiNo=m_AryRuiKei(0,i) Then

							w_RuiKekka	= m_AryRuiKei(1,i)	'//欠課数
							w_RuiTikoku = m_AryRuiKei(2,i)	'//遅刻数
							
							'//累積の表示法が累積の場合
							'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '試験毎
							'//Public Const C_K_KEKKA_RUISEKI_KEI = 1	   '累積
							
							If cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) Then
								
								'//後期期末試験後の場合
								'If m_iSikenKbn = 0 Then
								'	w_iKbn = 4
								'Else
								'	w_iKbn = cint(m_iSikenKbn) - 1
								'End If
								
								'//前回の試験時の出欠を取得する
								w_sGakusekiNo = m_Rs_M("GAKUSEI")		'2001/12/17 Add
								'm_sSyubetu
								'Call gf_GetKekaChi(m_iSyoriNen,m_iShikenInsertType,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
								Call gf_GetKekaChi(m_iSyoriNen,m_sSyubetu,m_iShikenInsertType,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
								
								w_RuiKekka = cint(w_RuiKekka) + cint(gf_SetNull2Zero(p_iKekka))
								w_RuiTikoku = cint(w_RuiTikoku) + cint(gf_SetNull2Zero(p_iChikoku))

							End If

							Exit For
						End If
					Next

					%>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiTikoku <> 0, w_TukiTikoku, "　") %><br></td>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiKekka <> 0, w_TukiKekka, "　") %><br></td>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiTikoku <> 0, w_RuiTikoku, "　") %><br></td>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiKekka <> 0,  w_RuiKekka, "　")  %><br></td>
					</tr>
					<%m_Rs_M.MoveNext%>
				<%Loop%>
			<%End If%>
			<%
			'=================================================
			'//代替留学生累計の追加
			If m_bDaigae = True Then
				If m_Rs_D.EOF = False Then
					'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
					Call gs_cellPtn(w_Class) 
					
					Do Until m_Rs_D.EOF
					%>
						<tr>
						<%
						w_sGakusekiNo = m_Rs_D("GAKUSEKI")
						
						'//初期化
						
						w_TukiTikoku = 0   '//月遅刻
						w_TukiKekka  = 0   '//月欠課
						w_RuiTikoku  = 0   '//累計遅刻
						w_RuiKekka	 = 0   '//累計欠課
						
						'//月計を取得
						For i=0 To cInt(m_iTukiKeiCnt)
							If w_sGakusekiNo=m_AryTukiKei(0,i) Then
								w_TukiKekka  = m_AryTukiKei(1,i)	'//欠課数
								w_TukiTikoku = m_AryTukiKei(2,i)	'//遅刻数
								Exit For
							End If
						Next

						'//累計を取得
						For i=0 To m_iRuiKeiCnt
							If w_sGakusekiNo=m_AryRuiKei(0,i) Then
								w_RuiKekka	= m_AryRuiKei(1,i)	'//欠課数
								w_RuiTikoku = m_AryRuiKei(2,i)	'//遅刻数


								'// 累積の表示法が累積の場合
								'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '試験毎
								'//Public Const C_K_KEKKA_RUISEKI_KEI = 1	   '累積
								
								If m_iSyubetu = C_K_KEKKA_RUISEKI_KEI Then
									
									'//後期期末試験後の場合
									'If m_iSikenKbn = 0 Then
									'	w_iKbn = 4
									'Else
									'	w_iKbn = cint(m_iSikenKbn) - 1
									'End If
									
									'//前回の試験時の出欠を取得する
									w_sGakusekiNo = m_Rs_D("GAKUSEI")
									'Call gf_GetKekaChi(m_iSyoriNen,w_iKbn,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
									Call gf_GetKekaChi(m_iSyoriNen,m_sSyubetu,m_iShikenInsertType,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
									
									w_RuiKekka = cint(w_RuiKekka) + cint(gf_SetNull2Zero(p_iKekka))
									w_RuiTikoku = cint(w_RuiTikoku) + cint(gf_SetNull2Zero(p_iChikoku))
									
								End If

								Exit For
							End If
						Next
						%>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiTikoku <> 0, w_TukiTikoku, "　") %><br></td>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiKekka <> 0, w_TukiKekka, "　") %><br></td>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiTikoku <> 0, w_RuiTikoku, "　") %><br></td>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiKekka <> 0,  w_RuiKekka, "　")  %><br></td>
						</tr>
						<%m_Rs_D.MoveNext
					Loop
				End If

			End If
			'=================================================
			%>

			</table>

		</td>
		<tr><td height=10><br></td></tr>
		<tr>
		
		<% if w_bEndFLG = False Then %> 
			<td valign="bottom"  colspan=3 align="center">
				<input class=button type="button" onclick="javascript:f_Touroku();" value="　登　録　">
				&nbsp;&nbsp;&nbsp;
				<input class=button type="button" onclick="javascript:f_Cancel();" value="キャンセル">
			</td>
		<% else %>
			<td valign="bottom"  colspan=3 align="center">
				<input class=button type="button" onclick="javascript:f_Cancel();" value=" 戻　る ">
			</td>
		<% end If %>
		
		</tr>
		</table>


		<%Exit Do%>
	<%Loop%>

	<input type="hidden" name="JikanSU"   value="<%=m_iRsCnt%>">
	<input type="hidden" name="Tuki_Zenki_Start" value="<%=m_sZenki_Start%>">
	<input type="hidden" name="Tuki_Kouki_Start" value="<%=m_sKouki_Start%>">
	<input type="hidden" name="Tuki_Kouki_End"	 value="<%=m_sKouki_End%>">
	<INPUT TYPE=HIDDEN NAME="NENDO" 	value = "<%=m_iSyoriNen%>">
	<INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
	<INPUT TYPE=HIDDEN NAME="TUKI"		value = "<%=m_sTuki%>">
	<INPUT TYPE=HIDDEN NAME="GAKKI" 	value = "<%=m_sGakki%>">
	<INPUT TYPE=HIDDEN NAME="GAKUNEN"	value = "<%=m_sGakunen%>">
	<INPUT TYPE=HIDDEN NAME="CLASSNO"	value = "<%=m_sClassNo%>">
	<INPUT TYPE=HIDDEN NAME="KAMOKU_CD" value = "<%=m_sKamokuCd%>">
	<INPUT TYPE=HIDDEN NAME="SYUBETU"	value = "<%=m_sSyubetu%>">
	<INPUT TYPE=HIDDEN NAME="EndFLG"   value = "<%=w_bEndFLG%>">
	<INPUT TYPE=HIDDEN NAME="cboGakunenCd"   value = "<%=request("cboGakunenCd")%>">
	<INPUT TYPE=HIDDEN NAME="cboClassCd"   value = "<%=request("cboClassCd")%>">
	<INPUT TYPE=HIDDEN NAME="txtFromDate"   value = "<%=request("txtFromDate")%>">
	<INPUT TYPE=HIDDEN NAME="txtToDate"   value = "<%=request("txtToDate")%>">
	
	
	<INPUT TYPE=HIDDEN NAME="KAMOKU_NAME" value="<%=Request("KAMOKU_NAME")%>">
	<INPUT TYPE=HIDDEN NAME="CLASS_NAME"  value="<%=Request("CLASS_NAME")%>">

	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	[機能]	空白HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>授業出欠入力</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {
	}
	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" mothod="post">

	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>

	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub
%>
