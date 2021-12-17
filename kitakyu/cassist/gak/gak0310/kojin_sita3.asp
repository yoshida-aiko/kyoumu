<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_sita1.asp
' 機      能: 検索された学生の詳細を表示する(個人情報)
'-------------------------------------------------------------------------
' 引      数	Session("GAKUSEI_NO")  = 学生番号
'            	Session("HyoujiNendo") = 表示年度
'           
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg        'ｴﾗｰﾌﾗｸﾞ
	Public m_Rs				'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
	Public m_SEIBETU		'性別
	Public m_BLOOD			'血液型
	Public m_RH				'RH
	Public m_HOG_ZOKU		'保護者続柄
	Public m_HOS_ZOKU		'保証人続柄
	Public m_RYOSEI_KBN		'寮生区分
	Public m_RYUNEN_FLG		'進級区分

	Public m_HyoujiFlg		'表示ﾌﾗｸﾞ
	Public m_KakoRs			'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ(過去ｸﾗｽ)
	Public mHyoujiNendo		'表示年度

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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="学生情報検索結果"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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
		'//過去のクラスを取得
		w_iRet = f_GetDetailKakoClass()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'//表示項目を取得
		w_iRet = f_GetDetailGakunen()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//初期表示
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    '// 終了処理
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  過去のクラスを取得
'*  [引数]  なし
'*  [戻値]  0:正常終了	1:任意のエラー  99:システムエラー
'*  [説明]  
'********************************************************************************
Function f_GetDetailKakoClass()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKakoClass = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	T13.T13_NENDO, "
		w_sSql = w_sSql & " 	T13.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	T13.T13_CLASS "
		w_sSql = w_sSql & " FROM T13_GAKU_NEN T13 "
		w_sSql = w_sSql & " WHERE  "
		w_sSql = w_sSql & " 	    T13.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		'w_sSql = w_sSql & " 	AND T13.T13_RYUNEN_FLG <> 1"
		w_sSql = w_sSql & " ORDER BY T13.T13_NENDO DESC "

		iRet = gf_GetRecordset(m_KakoRs, w_sSql)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailKakoClass = 99
			Exit Do
		End If

		if m_KakoRs.Eof then
			msMsg = "学年情報取得時にエラーが発生しました"
			f_GetDetailKakoClass = 99
			Exit Do
		End if

		'//正常終了
		f_GetDetailKakoClass = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  クラス情報取得
'*  [引数]  なし
'*  [戻値]  クラス名称
'*  [説明]  
'********************************************************************************
Function f_GetClass(p_sCLASS,p_iGakunen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClass = ""

	Do 

		'// クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "   M05_CLASSMEI	 "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASSRYAKU	 "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_TANNIN	"	
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		'w_sSQL = w_sSQL & vbCrLf & "      M05_NENDO = " & request("selNendo")
		
		If request("selNendo") = "" Then
			w_sSQL = w_sSQL & vbCrLf & "      M05_NENDO = " & mHyoujiNendo
		Else
			w_sSQL = w_sSQL & vbCrLf & "      M05_NENDO = " & request("selNendo")
		End If

		w_sSQL = w_sSQL & vbCrLf & "  AND M05_GAKUNEN =" & p_iGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASSNO = '" & p_sCLASS & "'"
'response.write w_ssql
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			Exit Do
		End If

		If rs.EOF Then
			Exit Do
		End If

		f_GetClass = gf_HTMLTableSTR(rs("M05_CLASSMEI"))

		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  表示項目を取得
'*  [引数]  なし
'*  [戻値]  0:正常終了	1:任意のエラー  99:システムエラー
'*  [説明]  
'********************************************************************************
Function f_GetDetailGakunen()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	'// 表示する年度を決める
	wSelNendo = request("selNendo")
	if gf_IsNull(wSelNendo) then
		mHyoujiNendo = Session("HyoujiNendo")
	Else
		mHyoujiNendo = wSelNendo
	End if

	f_GetDetailGakunen = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & "		A.T13_NENDO, "				'処理年度
		w_sSql = w_sSql & "		A.T13_GAKUSEKI_NO, "		'学籍番号
		w_sSql = w_sSql & "		A.T13_GAKUNEN, "			'学年
		w_sSql = w_sSql & " 	E.M01_SYOBUNRUIMEI, "		'在籍区分　＊

		w_sSql = w_sSql & " 	E.M01_DAIBUNRUI_CD, "		'在籍区分　＊
		w_sSql = w_sSql & " 	E.M01_SYOBUNRUI_CD, "		'在籍区分　＊

		w_sSql = w_sSql & "		A.T13_GAKKA_CD, "			'学科CD
		w_sSql = w_sSql & " 	B.M02_GAKKAMEI, "			'学科名称　＊
		w_sSql = w_sSql & "		A.T13_COURCE_CD, "			'コースCD　＊
		w_sSql = w_sSql & "		A.T13_CLASS, "				'クラスCD　＊
		w_sSql = w_sSql & "		A.T13_SYUSEKI_NO1, "		'出席番号（学科）
		w_sSql = w_sSql & "		A.T13_SYUSEKI_NO2, "		'出席番号（クラス）
		w_sSql = w_sSql & "		A.T13_RYOSEI_KBN, "			'寮生区分　＊
		w_sSql = w_sSql & "		A.T13_RYUNEN_FLG, "			'留年区分　＊
		w_sSql = w_sSql & "		A.T13_CLUB_1, "				'クラブ１　＊
		w_sSql = w_sSql & "		A.T13_CLUB_1_NYUBI, "		'クラブ入日１
		w_sSql = w_sSql & "		A.T13_CLUB_2, "				'クラブ２　＊
		w_sSql = w_sSql & "		A.T13_CLUB_2_NYUBI, "		'クラブ入日２
		w_sSql = w_sSql & "		A.T13_TOKUKATU, "			'特別活動
		w_sSql = w_sSql & "		A.T13_TOKUKATU_DET, "		'特別活動詳細
		w_sSql = w_sSql & "		A.T13_NENSYOKEN, "			'指導上参考となる諸事項
		w_sSql = w_sSql & "		A.T13_NENSYOKEN2, "			'指導上参考となる諸事項2
		w_sSql = w_sSql & "		A.T13_NENSYOKEN3, "			'指導上参考となる諸事項3
		w_sSql = w_sSql & "		A.T13_SINTYO, "				'身長
		w_sSql = w_sSql & "		A.T13_TAIJYU, "				'体重
		w_sSql = w_sSql & "		A.T13_SEKIJI_TYUKAN_Z, "	'前期中間席次
		w_sSql = w_sSql & "		A.T13_SEKIJI_KIMATU_Z, " 	'前期期末席次
		w_sSql = w_sSql & "		A.T13_SEKIJI_TYUKAN_K, " 	'後期期末席次
		w_sSql = w_sSql & "		A.T13_SEKIJI, "				'学年末席次
		w_sSql = w_sSql & "		A.T13_NINZU_TYUKAN_Z, "		'前期中間クラス人数
		w_sSql = w_sSql & "		A.T13_NINZU_KIMATU_Z, "  	'前期期末クラス人数
		w_sSql = w_sSql & "		A.T13_NINZU_TYUKAN_K, "  	'後期中間クラス人数
		w_sSql = w_sSql & "		A.T13_CLASSNINZU, "			'学年末クラス人数
		w_sSql = w_sSql & "		A.T13_HEIKIN_TYUKAN_Z, " 	'前期中間平均点
		w_sSql = w_sSql & "		A.T13_HEIKIN_KIMATU_Z, " 	'前期期末平均点
		w_sSql = w_sSql & "		A.T13_HEIKIN_TYUKAN_K, " 	'後期中間平均点
		w_sSql = w_sSql & "		A.T13_HEIKIN_KIMATU_K, " 	'学年末平均点
		w_sSql = w_sSql & "		A.T13_SUMJYUGYO, "			'総授業日数
		w_sSql = w_sSql & "		A.T13_SUMSYUSSEKI, "		'出席日数
		w_sSql = w_sSql & "		A.T13_SUMRYUGAKU, "			'留学中の授業日数
		w_sSql = w_sSql & "		A.T13_KESSEKI_TYUKAN_Z, "	'前期中間欠席日数
		w_sSql = w_sSql & "		A.T13_KESSEKI_KIMATU_Z, "	'前期期末欠席日数
		w_sSql = w_sSql & "		A.T13_KESSEKI_TYUKAN_K, "	'期末中間欠席日数
		w_sSql = w_sSql & "		A.T13_SUMKESSEKI, "			'学年末欠席日数（総欠席日数）
		w_sSql = w_sSql & "		A.T13_KIBIKI_TYUKAN_Z, "	'前期中間忌引き日数
		w_sSql = w_sSql & "		A.T13_KIBIKI_KIMATU_Z, "	'前期期末忌引き日数
		w_sSql = w_sSql & "		A.T13_KIBIKI_TYUKAN_K, "	'後期中間期引き日数
		w_sSql = w_sSql & "		A.T13_SUMKIBTEI, "			'出席停止忌引き日数（総忌引き数）
		w_sSql = w_sSql & "		A.T13_NENBIKO "				'指導用参考となる諸事項
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	M02_GAKKA    B, "
		w_sSql = w_sSql & " 	M01_KUBUN E  "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	 A.T13_GAKKA_CD   = B.M02_GAKKA_CD(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = B.M02_NENDO(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = " & mHyoujiNendo
		w_sSql = w_sSql & "  AND A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql = w_sSql & " 	AND A.T13_NENDO		   = E.M01_NENDO "
		w_sSql = w_sSql & " 	AND E.M01_DAIBUNRUI_CD = " & C_ZAISEKI				'在籍区分
		w_sSql = w_sSql & " 	AND E.M01_SYOBUNRUI_CD = T13_ZAISEKI_KBN "				'在籍区分

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailGakunen = 99
			Exit Do
		End If

		'//寮生区分を取得
		if Not gf_GetKubunName(C_NYURYO,m_Rs("T13_RYOSEI_KBN"),Session("HyoujiNendo"),m_RYOSEI_KBN) then Exit Do

		'//

		'//進級区分を取得
		Select Case gf_SetNull2String(m_Rs("T13_RYUNEN_FLG"))
			Case "0"
				m_RYUNEN_FLG = " － "
			Case "1"
				m_RYUNEN_FLG = "留年"
			Case Else
				m_RYUNEN_FLG = " － "
		End Select

		'//正常終了
		f_GetDetailGakunen = 0
		Exit Do
	Loop

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
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & mHyoujiNendo
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
'*  [機能]  担任を取得する
'*  [引数]  p_sGAKKACd:学科CD
'*  [戻値]  f_GetTanninName：担任名
'*  [説明]  
'********************************************************************************
Function f_GetTanninName(p_sGAKKACd,p_iGAKUNEN)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetTanninName = ""
	w_sTanninName = ""

	Do

		'//学科CDが空の時
		If trim(gf_SetNull2String(p_sGAKKACd)) = "" Then
			Exit Do
		End If
		'//学年が空の時
		If trim(gf_SetNull2Zero(p_iGAKUNEN)) = 0 Then
			Exit Do
		End If

		'//担任取得
		w_sSql = "Select "
	    w_sSql = w_sSql & " M04_KYOKANMEI_SEI,"
	    w_sSql = w_sSql & " M04_KYOKANMEI_MEI "
	    w_sSql = w_sSql & " From"
	    w_sSql = w_sSql & " M05_CLASS,"
	    w_sSql = w_sSql & " M04_KYOKAN "
	    w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " M05_NENDO =" & mHyoujiNendo
		w_sSql = w_sSql & " And "
	    w_sSql = w_sSql & " M05_NENDO = M04_NENDO "
		w_sSql = w_sSql & " And "
	    'w_sSql = w_sSql & " M04_GAKKA_CD =" & p_sGAKKACd   '学科コード
	    w_sSql = w_sSql & " M05_CLASSNO =" & p_sGAKKACd   'クラスコード
	    w_sSql = w_sSql & " And "
		w_sSql = w_sSql & " M05_GAKUNEN =" & p_iGAKUNEN '学年
	    w_sSql = w_sSql & " And "
		w_sSql = w_sSql & " M05_TANNIN = M04_KYOKAN_CD " '教官
'response.write w_ssql
		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//部活名
			w_sTanninName = rs("M04_KYOKANMEI_SEI") & "  " & rs("M04_KYOKANMEI_MEI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetTanninName = w_sTanninName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  クラス委員を取得する
'*  [引数]  p_sGAKKACd:学科CD
'*  [戻値]  f_GetTanninName：担任名
'*  [説明]  
'********************************************************************************
Function f_GetIinName()

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	Dim w_sGakki
	Dim w_sZenki_Start
	Dim w_sKouki_Start
	Dim w_sKouki_End

	On Error Resume Next
	Err.Clear

	f_GetIinName = ""
	w_sIinName = ""
	w_sGakki = ""
	w_sZenki_Start = ""
	w_sKouki_Start = ""
	w_sKouki_End = ""

	Do
		'学期収得（現在の学期）
		Call gf_GetGakkiInfo(w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End)

		'//委員取得
		w_sSql = ""
		w_sSql = w_sSql & "SELECT "
		w_sSql = w_sSql & "M34_IIN_NAME "
		w_sSql = w_sSql & "FROM "
		w_sSql = w_sSql & "M34_IIN, "
		w_sSql = w_sSql & "T06_GAKU_IIN "
		w_sSql = w_sSql & "WHERE "
		w_sSql = w_sSql & "M34_NENDO =" & mHyoujiNendo
		w_sSql = w_sSql & " AND "
		w_sSql = w_sSql & "T06_NENDO = M34_NENDO "
		w_sSql = w_sSql & " AND "
		w_sSql = w_sSql & "M34_DAIBUN_CD = T06_DAIBUN_CD "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "M34_SYOBUN_CD = T06_SYOBUN_CD "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "T06_IIN_KBN=2 "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "T06_GAKKI_KBN = " & w_sGakki
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "M34_IIN_KBN = T06_IIN_KBN "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "T06_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//委員名
			w_sIinName = rs("M34_IIN_NAME")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetIinName = w_sIinName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  クラス委員を取得する
'*  [引数]  p_sGAKKACd:学科CD、p_sCourseCd:コースCD
'*  [戻値]  p_sCourseCd：コース名
'*  [説明]  
'********************************************************************************
Function f_GetCourseName(p_iNendo,p_iGakunen,p_sGakkaCd,p_sCourseCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_sName

	On Error Resume Next
	Err.Clear

	f_GetCourseName = ""
	w_sName = ""

	Do
		'//委員取得
		w_sSql = ""
		w_sSql = ""
		w_sSql = w_sSql & "SELECT "
		w_sSql = w_sSql & "M20_NENDO,"
		w_sSql = w_sSql & "M20_GAKKA_CD,"
		w_sSql = w_sSql & "M20_GAKUNEN,"
		w_sSql = w_sSql & "M20_COURSE_CD,"
		w_sSql = w_sSql & "M20_COURSEMEI,"
		w_sSql = w_sSql & "M20_COURSEMEI_EIGO,"
		w_sSql = w_sSql & "M20_COURSERYAKSYO,"
		w_sSql = w_sSql & "M20_COURSE_KIGO,"
		w_sSql = w_sSql & "M20_COURSE_TEIIN "
		w_sSql = w_sSql & "FROM "
		w_sSql = w_sSql & "M20_COURSE "
		w_sSql = w_sSql & "WHERE "
		w_sSql = w_sSql & "M20_NENDO = " & p_iNendo & " AND "
		w_sSql = w_sSql & "M20_GAKKA_CD = '" & p_sGakkaCd & "' AND "
		w_sSql = w_sSql & "M20_GAKUNEN = " & p_iGakunen & " AND "
		w_sSql = w_sSql & "M20_COURSE_CD = '" & p_sCourseCd & "'"

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//委員名
			w_sName = rs("M20_COURSEMEI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetCourseName = w_sName

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

	On Error Resume Next
	Err.Clear

	m_HyoujiFlg = 0 		'<!-- 表示フラグ（0:なし  1:あり）

	m_NENDO				= "" 	'処理年度
	m_GAKUSEKI_NO		= "" 	'学籍番号
	m_GAKUNEN			= "" 	'学年
	m_ZAISEKI_KBN		= "" 	'在籍区分　＊
	m_GAKKA_CD			= "" 	'学科CD　＊
	m_COURCE_CD			= "" 	'コースCD　＊
	m_CLASS				= "" 	'クラスCD　＊
	m_TANNIN			= "" 	'担任名　＊
	m_SYUSEKI_NO1		= "" 	'出席番号（学科）
	m_SYUSEKI_NO2		= "" 	'出席番号（クラス）
	'm_RYOSEI_KBN		= "" 	'寮生区分　＊
	m_RYUNEN_FLG		= "" 	'留年区分　＊
	m_IIN				= "" 	'クラス役員　＊
	m_CLUB_1			= "" 	'クラブ１　＊
	m_CLUB_1_NYUBI		= "" 	'クラブ入日１
	m_CLUB_2			= "" 	'クラブ２　＊
	m_CLUB_2_NYUBI		= "" 	'クラブ入日２
	m_TOKUKATU			= "" 	'特別活動
	m_TOKUKATU_DET		= "" 	'特別活動詳細
	m_NENSYOKEN			= ""	'指導上参考となる諸事項
	m_NENSYOKEN2		= ""	'指導上参考となる諸事項2
	m_NENSYOKEN3		= ""	'指導上参考となる諸事項3
	m_SINTYO			= "" 	'身長
	m_TAIJYU			= "" 	'体重
	m_SEKIJI_TYUKAN_Z	= "" 	'前期中間席次
	m_SEKIJI_KIMATU_Z 	= "" 	'前期期末席次
	m_SEKIJI_TYUKAN_K 	= "" 	'後期期末席次
	m_SEKIJI		 	= "" 	'学年末席次
	m_NINZU_TYUKAN_Z	= ""	'前期中間クラス人数
	m_NINZU_KIMATU_Z	= "" 	'前期期末クラス人数
	m_NINZU_TYUKAN_K  	= "" 	'後期中間クラス人数
	m_CLASSNINZU		= "" 	'学年末クラス人数
	m_HEIKIN_TYUKAN_Z 	= "" 	'前期中間平均点
	m_HEIKIN_KIMATU_Z 	= "" 	'前期期末平均点
	m_HEIKIN_TYUKAN_K 	= "" 	'後期中間平均点
	m_HEIKIN_KIMATU_K 	= "" 	'学年末平均点
	m_SUMJYUGYO			= "" 	'総授業日数
	m_SUMSYUSSEKI		= "" 	'出席日数
	m_SUMRYUGAKU		= "" 	'留学中の授業日数
	m_KESSEKI_TYUKAN_Z	= "" 	'前期中間欠席日数
	m_KESSEKI_KIMATU_Z	= ""	'前期期末欠席日数
	m_KESSEKI_TYUKAN_K	= ""	'期末中間欠席日数
	m_SUMKESSEKI		= "" 	'学年末欠席日数（総欠席日数）
	m_KIBIKI_TYUKAN_Z	= ""	'前期中間忌引き日数
	m_KIBIKI_KIMATU_Z	= ""	'前期期末忌引き日数
	m_KIBIKI_TYUKAN_K	= ""	'後期中間期引き日数
	m_SUMKIBTEI			= ""	'出席停止忌引き日数（総忌引き数）
'					    = "" 	'授業料免除
'					　  = "" 	'奨学金
	m_NENBIKO		    = "" 	'指導用参考となる諸事項

	if Not m_Rs.Eof Then
		m_NENDO			= m_Rs("T13_NENDO")
		m_GAKUSEKI_NO	= m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN		= m_Rs("T13_GAKUNEN")
		m_ZAISEKI_KBN	= m_Rs("M01_SYOBUNRUIMEI")
		m_GAKKAMEI	 	= m_Rs("M02_GAKKAMEI")
		m_COURCE_CD		= m_Rs("T13_COURCE_CD")
		m_COURCEMEI		= f_GetCourseName(gf_SetNull2Zero(m_Rs("T13_NENDO")),gf_SetNull2Zero(m_Rs("T13_GAKUNEN")),gf_SetNull2String(m_Rs("T13_GAKKA_CD")),gf_SetNull2String(m_Rs("T13_COURCE_CD")))
		m_CLASS			= m_Rs("T13_CLASS")
		m_TANNIN		= f_GetTanninName(gf_SetNull2String(m_Rs("T13_CLASS")),gf_SetNull2Zero(m_Rs("T13_GAKUNEN")))
		m_SYUSEKI_NO1	= m_Rs("T13_SYUSEKI_NO1")
		m_SYUSEKI_NO2	= m_Rs("T13_SYUSEKI_NO2")
		m_IIN			= f_GetIinName
		'm_RYOSEI_KBN	= m_Rs("T13_RYOSEI_KBN")

		If gf_SetNull2String(m_Rs("T13_RYUNEN_FLG")) = "1" Then
			m_RYUNEN_FLG	= "留年"
		Else
			m_RYUNEN_FLG	= " － "
		End If

		m_CLUB_1		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_1")))
		m_CLUB_1_NYUBI	= m_Rs("T13_CLUB_1_NYUBI")
		m_CLUB_2		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_2")))
		m_CLUB_2_NYUBI	= m_Rs("T13_CLUB_2_NYUBI")
		m_TOKUKATU		= m_Rs("T13_TOKUKATU")
		m_TOKUKATU_DET	= m_Rs("T13_TOKUKATU_DET")
		m_NENSYOKEN		= m_Rs("T13_NENSYOKEN")
		m_NENSYOKEN2	= m_Rs("T13_NENSYOKEN2")
		m_NENSYOKEN3	= m_Rs("T13_NENSYOKEN3")
		m_SINTYO		= m_Rs("T13_SINTYO")
		m_TAIJYU		= m_Rs("T13_TAIJYU")
		m_SEKIJI_TYUKAN_Z	= m_Rs("T13_SEKIJI_TYUKAN_Z")
		m_SEKIJI_KIMATU_Z	= m_Rs("T13_SEKIJI_KIMATU_Z")
		m_SEKIJI_TYUKAN_K	= m_Rs("T13_SEKIJI_TYUKAN_K")
		m_SEKIJI		 	= m_Rs("T13_SEKIJI")
		m_NINZU_TYUKAN_Z	= m_Rs("T13_NINZU_TYUKAN_Z")
		m_NINZU_KIMATU_Z	= m_Rs("T13_NINZU_KIMATU_Z")
		m_NINZU_TYUKAN_K	= m_Rs("T13_NINZU_TYUKAN_K")
		m_CLASSNINZU		= m_Rs("T13_CLASSNINZU")
		m_HEIKIN_TYUKAN_Z	= m_Rs("T13_HEIKIN_TYUKAN_Z")
		m_HEIKIN_KIMATU_Z	= m_Rs("T13_HEIKIN_KIMATU_Z")
		m_HEIKIN_TYUKAN_K	= m_Rs("T13_HEIKIN_TYUKAN_K")
		m_HEIKIN_KIMATU_K	= m_Rs("T13_HEIKIN_KIMATU_K")
		m_SUMJYUGYO			= m_Rs("T13_SUMJYUGYO")
		m_SUMSYUSSEKI		= m_Rs("T13_SUMSYUSSEKI")
		m_SUMRYUGAKU		= m_Rs("T13_SUMRYUGAKU")
		m_KESSEKI_TYUKAN_Z	= m_Rs("T13_KESSEKI_TYUKAN_Z")
		m_KESSEKI_KIMATU_Z	= m_Rs("T13_KESSEKI_KIMATU_Z")
		m_KESSEKI_TYUKAN_K	= m_Rs("T13_KESSEKI_TYUKAN_K")
		m_SUMKESSEKI		= m_Rs("T13_SUMKESSEKI")
		m_KIBIKI_TYUKAN_Z	= m_Rs("T13_KIBIKI_TYUKAN_Z")
		m_KIBIKI_KIMATU_Z	= m_Rs("T13_KIBIKI_KIMATU_Z")
		m_KIBIKI_TYUKAN_K	= m_Rs("T13_KIBIKI_TYUKAN_K")
		m_SUMKIBTEI	= m_Rs("T13_SUMKIBTEI")
		m_NENBIKO	= m_Rs("T13_NENBIKO")
        
	End if

%>

	<html>
	<head>
	<title>学籍データ参照</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	<script language="javascript">
	<!--
		//**************************************
		//*   年度ｾﾚｸﾄﾎﾞｯｸｽが変更されたとき
		//**************************************
		function jf_ChangSelect(){

			document.frm.submit();

		}

	//-->
	</script>
	</head>

	<body>
	<form action="kojin_sita3.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">●基本情報</a></td>
			<td nowrap><a href="kojin_sita1.asp">●個人情報</a></td>
			<td nowrap><a href="kojin_sita2.asp">●入学情報</a></td>
			<td nowrap><b>●学年情報</b></td>
			<td nowrap><a href="kojin_sita4.asp">●その他予備情報</a></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td colspan="3">
				<span class="msg"><font size="2">※ 処理年度を変更すると、過去の学年情報を見ることができます<BR></font></span>
			</td>
		</tr>
		<tr>
			<td valign="top" align="left">

				<table class="disp" border="1" width="240">
						<tr>
							<td class="disph" width="80">処理年度</td>
							<td class="disp"><select name="selNendo" onChange="jf_ChangSelect();">
												<% do until m_KakoRs.Eof 
													wSelected = ""
													if Cint(mHyoujiNendo) = Cint(m_KakoRs("T13_NENDO")) then
														wSelected = "selected"
													End if
													%>
													<option value="<%=m_KakoRs("T13_NENDO")%>" <%=wSelected%>><%=m_KakoRs("T13_NENDO")%>年度
												<% m_KakoRs.MoveNext : Loop %>
											</select></td>
						</tr>
<!--
						<tr>
							<td class="disph" width="80" height="16">処理年度</td>
							<td class="disp"><%= m_NENDO %>&nbsp</td>
						</tr>
-->

					<% if gf_empItem(C_T13_GAKUSEKI_NO) then %>
						<tr>
							<td class="disph" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_1NEN)%></td>
							<td class="disp"><%= m_GAKUSEKI_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKUNEN) then %>
						<tr>
							<td class="disph" height="16">学　　年</td>
							<td class="disp"><%= m_GAKUNEN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_ZAISEKI_KBN) then %>
						<tr>
							<td class="disph" height="16">在籍区分</td>
							<td class="disp"><%= m_ZAISEKI_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKKA_CD) then %>
						<tr>
							<td class="disph" height="16">所属学科</td>
							<td class="disp"><%= m_GAKKAMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_COURCE_CD) then %>
						<tr>
							<td class="disph" height="16">コース</td>
							<td class="disp"><%= m_COURCEMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLASS) then %>
						<tr>
							<td class="disph" height="16">クラス</td>
							<td class="disp"><%= f_GetClass(m_CLASS,m_GAKUNEN) %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T_TANNIN) then %>
						<tr>
							<td class="disph" height="16">担任名</td>
							<td class="disp"><%= m_TANNIN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO1) then %>
						<tr>
							<td class="disph" height="16">出席番号(学科)</td>
							<td class="disp"><%= m_SYUSEKI_NO1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO2) then %>
						<tr>
							<td class="disph" height="16">出席番号(クラス)</td>
							<td class="disp"><%= m_SYUSEKI_NO2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYOSEI_KBN) then %>
						<tr>
							<td class="disph" height="16">寮生区分</td>
							<td class="disp"><%= m_RYOSEI_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYUNEN_FLG) then %>
						<tr>
							<td class="disph" height="16">進級区分</td>
							<td class="disp"><%= m_RYUNEN_FLG %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T_CLASSIIN) then %>
						<tr>
							<td class="disph" height="16">クラス役員</td>
							<td class="disp"><%= m_IIN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_1) then %>
						<tr>
							<td class="disph" height="16">クラブ活動１</td>
							<td class="disp"><%= m_CLUB_1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_1_NYUBI) then %>
						<tr>
							<td class="disph" height="16">入部日１</td>
							<td class="disp"><%= m_CLUB_1_NYUBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_2) then %>
						<tr>
							<td class="disph" height="16">クラブ活動２</td>
							<td class="disp"><%= m_CLUB_2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_2_NYUBI) then %>
						<tr>
							<td class="disph" height="16">入部日２</td>
							<td class="disp"><%= m_CLUB_2_NYUBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TOKUKATU) then %>
						<tr>
							<td class="disph" height="16">特別活動</td>
							<td class="disp"><%= m_TOKUKATU %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TOKUKATU_DET) then %>
						<tr>
							<td class="disph" height="16">特別活動詳細</td>
							<td class="disp"><%= m_TOKUKATU_DET %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SINTYO) then %>
						<tr>
							<td class="disph" height="16">身　　長</td>
							<td class="disp"><%= m_SINTYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TAIJYU) then %>
						<tr>
							<td class="disph" height="16">体　　重</td>
							<td class="disp"><%= m_TAIJYU %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_SEKIJI_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期中間席次</td>
							<td class="disp"><%= m_SEKIJI_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SEKIJI_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期期末席次</td>
							<td class="disp"><%= m_SEKIJI_KIMATU_Z  %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SEKIJI_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">後期中間席次</td>
							<td class="disp"><%= m_SEKIJI_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SEKIJI) then %>
						<tr>
							<td class="disph" width="140" height="16">学年末席次</td>
							<td class="disp"><%= m_SEKIJI %>&nbsp</td>
						</tr>
					<% End if %>
			</table>
			<br>

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NINZU_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期中間クラス人数</td>
							<td class="disp"><%= m_NINZU_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_NINZU_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期期末クラス人数</td>
							<td class="disp"><%= m_NINZU_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">後期中間クラス人数</td>
							<td class="disp"><%= m_NINZU_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_KIMATU_K) then %>
						<tr>
							<td class="disph" width="140" height="16">学年末クラス人数</td>
							<td class="disp"><%= m_CLASSNINZU %>&nbsp</td>
						</tr>
					<% End if %>
			</table>
			<br>

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_HEIKIN_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期中間平均点</td>
							<td class="disp"><%= m_HEIKIN_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期期末平均点</td>
							<td class="disp"><%= m_HEIKIN_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">後期中間平均点</td>
							<td class="disp"><%= m_HEIKIN_TYUKAN_K  %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_KIMATU_K) then %>
						<tr>
							<td class="disph" width="140" height="16">学年末平均点</td>
							<td class="disp"><%= m_HEIKIN_KIMATU_K  %>&nbsp</td>
						</tr>
					<% End if %>
			</table>
			<br>

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_SUMJYUGYO) then %>
						<tr>
							<td class="disph" width="140" height="16">総 授 業 日 数</td>
							<td class="disp"><%= m_SUMJYUGYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMSYUSSEKI) then %>
						<tr>
							<td class="disph" width="140" height="16">出 席 日 数</td>
							<td class="disp"><%= m_SUMSYUSSEKI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMRYUGAKU) then %>
						<tr>
							<td class="disph" width="140" height="16">留学中の授業日数</td>
							<td class="disp"><%= m_SUMRYUGAKU %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KESSEKI_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期中間欠席日数</td>
							<td class="disp"><%= m_KESSEKI_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KESSEKI_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期期末欠席日数</td>
							<td class="disp"><%= m_KESSEKI_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KESSEKI_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">後期中間欠席日数</td>
							<td class="disp"><%= m_KESSEKI_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMKESSEKI) then %>
						<tr>
							<td class="disph" width="140" height="16">学年末欠席日数</td>
							<td class="disp"><%= m_SUMKESSEKI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KIBIKI_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期中間忌引日数</td>
							<td class="disp"><%= m_KIBIKI_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>					
					<% if gf_empItem(C_T13_KIBIKI_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">前期期末忌引日数</td>
							<td class="disp"><%= m_KIBIKI_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KIBIKI_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">後期中間忌引日数</td>
							<td class="disp"><%= m_KIBIKI_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMKIBTEI) then %>
						<tr>
							<td class="disph" width="140" height="16">出席停止・忌引日数</td>
							<td class="disp"><%= m_SUMKIBTEI %>&nbsp</td>
						</tr>
					<% End if %>
				</table>
			</td>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NENSYOKEN) then %>
						<tr><td class="disph" width="220" height="16">指導上参考となる諸事項</td></tr>
						<tr><td class="disph" width="220" height="16">(1)学習における特徴等<br>(2)行動の特徴、特技等</td></tr>
						<tr><td class="disp" valign="top" height="60"><%= m_NENSYOKEN %></td></tr>
					<% End if %>
					<% if gf_empItem(C_T13_NENSYOKEN2) then %>
						<tr><td class="disph" width="220" height="16">(3)部活動、ボランティア活動等<br>(4)取得資格、検定等</td></tr>
						<tr><td class="disp" valign="top" height="60"><%= m_NENSYOKEN2 %></td></tr>
					<% End if %>
					<% if gf_empItem(C_T13_NENSYOKEN3) then %>
						<tr><td class="disph" width="220" height="16">(5)その他</td></tr>
						<tr><td class="disp" valign="top" height="60"><%= m_NENSYOKEN3 %></td></tr>
					<% End if %>
				</table>

			</td>
		</tr>
	</table>

	<% if m_HyoujiFlg = 0 then %>
		<BR>
		表示できるデータがありません<BR>
		<BR>
	<% End if %>

	<BR>
	<input type="button" class="button" value="　閉じる　" onClick="parent.window.close();">

	</div>
	</form>
	</body>
	</html>
<% End Sub %>