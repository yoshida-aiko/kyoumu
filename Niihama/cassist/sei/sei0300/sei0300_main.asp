<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人別成績一覧
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0300/sei0300_top.asp
' 機      能: フレームページ 成績一覧の登録を行う
'-------------------------------------------------------------------------
' 引      数:NENDO			:年度
'            txtSikenKBN	:試験区分
'            txtGakuNo		:学年
'            txtClassNo		:クラスNO
'            txtGakusei		:学生NO
'            txtBeforGakuNo	:前の学籍NO
'            txtAfterGakuNo	:後の学籍NO
' 引      渡:
'            txtSikenKBN	:試験区分
'            txtGakuNo		:学年
'            txtClassNo		:クラスNO
'            txtGakusei		:学生NO
'            txtBeforGakuNo	:前の学籍NO
'            txtAfterGakuNo	:後の学籍NO
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/09/04 伊藤公子
' 変      更: 2003/10/24 高田：福島高専の場合、授業時間数・欠課時数・遅刻回数を累積して表示
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="./sei0300_com.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
	Public CONST C_KETTEN_LIMIT = 60
	Public CONST C_KENGEN_SEI0300_FULL = "FULL"	'//アクセス権限FULL
	Public CONST C_KENGEN_SEI0300_TAN = "TAN"	'//アクセス権限担任
	Public CONST C_KENGEN_SEI0300_GAK = "GAK"	'//アクセス権限学科

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	Public  m_iSyoriNen		'//処理年度

	'//メインの変数
	Public  m_sSikenKBN		'//試験区分
	Public  m_sGakunen		'//学年
	Public  m_sClassNo		'//クラス
	Public  m_sGakkaNo		'//学科NO
	Public  m_sGakusei		'//学生NO

	Public  m_sName			'//生徒名称
	Public  m_sGakusekiNo	'//学籍NO
	Public  m_sBeforGakuNo	'//学籍NOが前の学生
	Public  m_sAfterGakuNo	'//学籍NOがあとの学生

	Public  m_iMemberCnt	'//クラス人数
	Public  m_iSikiji		'//席次
	Public  m_iSyoken		'//所見
	Public  m_iAverage		'//平均点
	
	Public  m_AryResult()	'//成績データ格納配列
	Public  m_iCnt			'//成績データ件数
	Public  m_sKengen
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

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="成績一覧"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_parent"

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

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

		'//権限チェック
		w_iRet = f_CheckKengen(w_sKengen)
		If w_iRet <> 0 Then
            m_bErrFlg = True
			m_sErrMsg = "参照権限がありません。"
			Exit Do
		End If

		'//権限が担任の場合は担任クラス情報を取得する
		If w_sKengen = C_KENGEN_SEI0300_TAN Then

			'//担任クラス情報取得
			'//情報が取得できない場合は担任クラスが無い為、参照不可とする
			w_iRet = f_GetClassInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "参照権限がありません。"
				Exit Do
			End If

		ElseIf w_sKengen = C_KENGEN_SEI0300_GAK Then

			'//学科情報取得
			'//情報が取得できない場合は学科が無い為、参照不可とする
			w_iRet = f_GetGakkaInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "参照権限がありません。"
				Exit Do
			End If

		End If

		'//生徒情報取得
		w_iRet = f_GetGakuseiData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "生徒情報が取得できませんでした。"
			Exit Do
		End If

		'//席次、所見を取得
		w_iRet = f_GetGakuseiInfo()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "生徒情報が取得できませんでした。"
			Exit Do
		End If

		'//成績データ取得
		w_iRet = f_GetResultData()
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
		'w_sMsg = gf_GetErrMsg()
		'Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

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

    m_iSyoriNen = ""
	m_sSikenKBN = ""
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sGakkaNo  = ""
    m_sGakusei  = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

	m_iSyoriNen = Session("NENDO")
	m_sSikenKBN = Request("txtSikenKBN")
	m_sGakunen  = Request("txtGakuNo")
	m_sClassNo  = Request("txtClassNo")
	m_sGakkaNo  = Request("txtGakkaNo")
	m_sGakusei  = Request("txtGakusei")

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_sSikenKBN = " & m_sSikenKBN & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sGakkaNo  = " & m_sGakkaNo  & "<br>"
    response.write "m_sGakusei  = " & m_sGakusei  & "<br>"

End Sub

'********************************************************************************
'*	[機能]	権限チェック
'*	[引数]	なし
'*	[戻値]	w_sKengen
'*	[説明]	ログインUSERの処理レベルにより、参照可不可の判断をする
'*			①FULLアクセス権限保持者は、全ての生徒の成績情報を参照できる
'*			②担任アクセス権限保持者は、受け持ちクラス生徒の成績情報を参照できる
'*			③上記以外のUSERは参照権限なし
'********************************************************************************
Function f_CheckKengen(p_sKengen)
    Dim w_iRet
    Dim w_sSQL
	 Dim rs

	 On Error Resume Next
	 Err.Clear

	 f_CheckKengen = 1

	 Do

		'T51より権限情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID IN ('SEI0300','SEI0301','SEI0302')"
		w_sSql = w_sSql & vbCrLf & "  AND T51_SYORI_LEVEL.T51_LEVEL" & Session("LEVEL") & " = 1"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			m_sErrMsg = Err.description
			f_CheckKengen = 99
			Exit Do
		End If

		If rs.EOF Then
			m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else
			Select Case rs("T51_ID")
				Case "SEI0300"	'//フルアクセス権限あり
					p_sKengen = C_KENGEN_SEI0300_FULL
				Case "SEI0301"	'//担任権限有り
					p_sKengen = C_KENGEN_SEI0300_TAN
				Case "SEI0302"	'//学科権限有り
					p_sKengen = C_KENGEN_SEI0300_GAK
			End Select

		End If

		f_CheckKengen = 0
		Exit Do
	 Loop


	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  権限チェック（担任クラス情報取得）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  ○担任アクセス権限が設定されているUSERでも、実際に担任クラスを
'*			受け持っていない場合には参照不可とする
'********************************************************************************
Function f_GetClassInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassInfo = 1
	p_sKengen = ""

	Do 

		'// 担任クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS.M05_CLASSNO "
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & session("KYOKAN_CD") & "'"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetClassInfo = 99
			Exit Do
		End If

		If rs.EOF Then
			'//クラス情報が取得できないとき
            m_sErrMsg = "参照権限がありません。"
			Exit Do
		End If

		f_GetClassInfo = 0
		p_sKengen = C_KENGEN_SEI0300_TAN
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  権限チェック（ユーザ学科情報取得）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  ○担任アクセス権限が設定されているUSERでも、実際に担任クラスを
'*			受け持っていない場合には参照不可とする
'********************************************************************************
Function f_GetGakkaInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetGakkaInfo = 1
	p_sKengen = ""

	Do 

		'// 担任クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M04_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & " FROM M04_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M04_KYOKAN_CD='" & session("KYOKAN_CD") & "'"
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetGakkaInfo = 99
			Exit Do
		End If
		If rs.EOF Then
			'//クラス情報が取得できないとき
            m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else
			p_sKengen = C_KENGEN_SEI0300_GAK 
'			m_sGakkaNo  = rs("M04_GAKKA_CD")
'			m_sGakkaMei = rs("M02_GAKKAMEI")

			'//権限が担任の場合は、担任クラス以外は選択できない
'			m_sGakuNoOption = " DISABLED "
'			m_sClassNoOption = " DISABLED "
		End If

		f_GetGakkaInfo = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

Function f_GetGakuseiData()
'********************************************************************************
'*  [機能]  生徒データを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim i
i = 1

	f_GetGakuseiData = 1

	Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT  "
		w_sSQL = w_sSQL & vbCrLf & "     A.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & "    ,A.T11_SIMEI "
		w_sSQL = w_sSQL & vbCrLf & "    ,B.T13_GAKUSEKI_NO "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "     T11_GAKUSEKI A,T13_GAKU_NEN B "
		w_sSQL = w_sSQL & vbCrLf & " WHERE"
		w_sSQL = w_sSQL & vbCrLf & "     B.T13_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_GAKUNEN = " & m_sGakunen
	If m_sKengen <> C_KENGEN_SEI0300_GAK then
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_CLASS = " & m_sClassNo
	Else
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_Gakka_CD = " & m_sGakkaNo
	End If
		w_sSQL = w_sSQL & vbCrLf & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
'		w_sSQL = w_sSQL & vbCrLf & " AND A.T11_NYUNENDO = B.T13_NENDO - B.T13_GAKUNEN + 1 "
		'//現在在学中の生徒のみ表示対象とする
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY B.T13_GAKUSEKI_NO "

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
	        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_GetGakuseiData = 99
			Exit do 
	    End If

		w_rCnt=cint(gf_GetRsCount(w_Rs))

		'//配列の作成
		w_Rs.MoveFirst
		Do Until w_Rs.EOF

			ReDim Preserve w_sGakuseiAry(i)
			w_sGakuseiAry(i) = w_Rs("T11_GAKUSEI_NO")
			i = i + 1

			If w_Rs("T11_GAKUSEI_NO") = m_sGakusei Then
				'//生徒名称を取得＆セット
				m_sName = w_Rs("T11_SIMEI")
				m_sGakusekiNo = w_Rs("T13_GAKUSEKI_NO")
			End If

			w_Rs.MoveNext

		Loop

		For i = 1 to w_rCnt

			If w_sGakuseiAry(i) = m_sGakusei Then

				'//学籍NOが前の生徒、後の生徒を取得＆セット
				If i <= 1 Then
					m_sAfterGakuNo = w_sGakuseiAry(i+1)
					Exit For
				End If

				If i = w_rCnt Then
					m_sBeforGakuNo = w_sGakuseiAry(i-1)
					Exit For
				End If

				m_sAfterGakuNo = w_sGakuseiAry(i+1)
				m_sBeforGakuNo = w_sGakuseiAry(i-1)

				Exit For
			End If

		Next

		f_GetGakuseiData = 0
		Exit Do
	Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  時限情報の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetResultData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_bSelect
    dim m_SchoolNo

    On Error Resume Next
    Err.Clear

    f_GetResultData = 1

		'学校番号を取得　2003.10.24 ins
		if Not gf_GetGakkoNO(m_SchoolNo) then
	        m_bErrFlg = True
			m_sErrMsg = "学校番号の取得に失敗しました。"
			Exit function
		end if
        m_SchoolNo = CSTR(m_SchoolNo)
'response.write m_SchoolNo
    Do

		'==============================
		'//試験の開始日、終了日を取得
		'==============================
		w_bRet = gf_GetKaisiSyuryo(cint(gf_SetNull2Zero(m_sSikenKBN)),cint(m_sGakunen), w_sKaisibi, w_sSyuryobi)

		If w_bRet <> True Then
			'//開始日、終了日の取得失敗
		    f_GetResultData = 99
			m_sErrMsg = "試験日程の登録を行ってください。"
			Exit Do 
		End If

		'==============================
		'//成績情報取得
		'==============================
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T16_KAMOKU_CD "
		w_sSql = w_sSql & vbCrLf & "  ,T16_KAMOKUMEI "
		w_sSql = w_sSql & vbCrLf & "  ,T16_HAITOTANI "
		w_sSql = w_sSql & vbCrLf & "  ,T16_KAISETU "
		w_sSql = w_sSql & vbCrLf & "  ,T16_HISSEN_KBN "
		w_sSql = w_sSql & vbCrLf & "  ,T16_SELECT_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T16_KAISETU "    '2003.10.28ins
		w_sSql = w_sSql & vbCrLf & "  ,M100_SEISEKI_INP "    '2003.10.28ins

		'//INS 2008/03/10 文字入力に対応
        if m_SchoolNo <> "11" then  '福島高専以外の場合
		'//試験区分により場合わけ
		Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

			Case C_SIKEN_ZEN_TYU    '前期中間試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_Z   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_Z AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_Z AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_ZEN_KIM    '前期期末試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_Z   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_Z AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_Z AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_TYU    '後期中間試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_K   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_K AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_K AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_K AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_KIM    '後期期末試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_K   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_K AS CHIKAI"
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_K AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_K AS HYOKA "	'INS 2008/03/10
			Case Else
				'//システムエラー
	            m_sErrMsg = "試験情報がありません。"
				Exit Do
		End Select
        END IF
        '================================================================================
        if m_SchoolNo = "11" then  '福島高専の場合
		'//試験区分により場合わけ
		Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

			Case C_SIKEN_ZEN_TYU    '前期中間試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_ZEN_KIM    '前期期末試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_TYU    '後期中間試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_K AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_KIM    '後期期末試験
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_K AS HYOKA "	'INS 2008/03/10

			Case Else
				'//システムエラー
	            m_sErrMsg = "試験情報がありません。"
				Exit Do
		End Select
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_Z   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_Z AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_Z AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_Z   AS KEKA2 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_Z AS CHIKAI2 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_Z AS JIKAN2 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_K   AS KEKA3 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_K AS CHIKAI3 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_K AS JIKAN3 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_K   AS KEKA4 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_K AS CHIKAI4 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_K AS JIKAN4 "
        END IF
        '================================================================================
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T16_RISYU_KOJIN"

'INS 2008/03/10
		w_sSql = w_sSql & vbCrLf & "  ,M03_KAMOKU"
		w_sSql = w_sSql & vbCrLf & "  ,M100_KAMOKU_ZOKUSEI"
'INS END 2008/03/10

		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T16_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND T16_GAKUSEI_NO='" & m_sGakusei & "'"

'INS 2008/03/10
		w_sSql = w_sSql & vbCrLf & "  AND T16_NENDO = M03_NENDO "
		w_sSql = w_sSql & vbCrLf & "  AND T16_KAMOKU_CD = M03_KAMOKU_CD "
		w_sSql = w_sSql & vbCrLf & "  AND M03_NENDO = M100_NENDO "
		w_sSql = w_sSql & vbCrLf & "  AND M03_ZOKUSEI_CD = M100_ZOKUSEI_CD "
		w_sSql = w_sSql & vbCrLf & "  AND M100_KAMOKUBUNRUI = '01' "
'INS END 2008/03/10

		w_sSql = w_sSql & vbCrLf & " ORDER BY T16_SEQ_NO"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			m_sErrMsg = "成績情報が取得できませんでした。"
            f_GetResultData = 99
            Exit Do
        End If

		'==============================
		'//成績情報を配列に格納する
		'==============================

		i = 0
		Do Until w_Rs.EOF

			'開設していない科目は表示しない
			do 
                '2003.10.28 upd_s

				'If f_GetKaisetu(gf_SetNull2String(w_Rs("T16_KAMOKU_CD"))) = False Then
				'	Exit Do
				'End If
		        if m_SchoolNo = "11" or m_SchoolNo = "55" then  '福島高専の場合
					Select Case cint(gf_SetNull2Zero(m_sSikenKBN))
						Case C_SIKEN_ZEN_TYU ,C_SIKEN_ZEN_KIM   '前期中間試験,前期期末試験

			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU")) = "1")) THEN
	                           Exit Do
							End If
						Case C_SIKEN_KOU_TYU    '後期中間試験
			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "2"))) THEN
								Exit Do
							End If
						Case C_SIKEN_KOU_KIM	'学年末は全て表示する
			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "1")) OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "2"))) THEN
								Exit Do
							End If
						Case Else
					End Select
				Else
	                '2003.10.28 upd_e
					Select Case cint(gf_SetNull2Zero(m_sSikenKBN))
						Case C_SIKEN_ZEN_TYU ,C_SIKEN_ZEN_KIM   '前期中間試験,前期期末試験

			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU")) = "1")) THEN
	                           Exit Do
							End If
						Case C_SIKEN_KOU_TYU,C_SIKEN_KOU_KIM    '後期中間試験,後期期末試験
			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "2"))) THEN
								Exit Do
							End If
						Case Else
					End Select

				END IF

				w_Rs.MoveNext

				If w_Rs.EOF = True Then 
					Exit Do
				End IF
			Loop

			If w_Rs.EOF = True Then 
				Exit Do
			End IF

			'//選択ﾌﾗｸﾞ初期化
			w_bSelect = False


			'//科目が必修の場合(C_HISSEN_HIS = 1 :必修)
			If cint(gf_SetNull2Zero(w_Rs("T16_HISSEN_KBN"))) = C_HISSEN_HIS Then
				w_bSelect = True
			Else

				'//科目が選択科目の場合(C_HISSEN_SEN = 2 : 選択)
				If cint(gf_SetNull2Zero(w_Rs("T16_HISSEN_KBN"))) = C_HISSEN_SEN Then

					'//選択科目を選択している場合(C_SENTAKU_YES = 1：選択する)
					If cint(gf_SetNull2Zero(w_Rs("T16_SELECT_FLG"))) = C_SENTAKU_YES Then
						w_bSelect = True
					Else
						w_bSelect = False
					End If

				End If

			End If

			If w_bSelect = True Then
			'	Redim Preserve m_AryResult(5,i)
				Redim Preserve m_AryResult(6,i)	'UPDATE 2008/03/10

				'//初期化
				m_AryResult(0,i) = ""
				m_AryResult(1,i) = ""
				m_AryResult(2,i) = ""
				m_AryResult(3,i) = ""
				m_AryResult(4,i) = ""
				m_AryResult(5,i) = ""

				m_AryResult(6,i) = ""

				m_AryResult(0,i) = w_Rs("T16_KAMOKUMEI")	'//科目名称
				m_AryResult(1,i) = w_Rs("T16_HAITOTANI")	'//配当単位

				'INS 2008/03/10
				m_AryResult(6,i) = w_Rs("M100_SEISEKI_INP")			'//遅刻回数

				IF Cint(m_AryResult(6,i)) = 0 then
					m_AryResult(2,i) = w_Rs("HTEN")			'//評価点
				else
					m_AryResult(2,i) = w_Rs("HYOKA")		'//評価
				end if
				'INS END 2008/03/10


               m_AryResult(3,i) = w_Rs("JIKAN")			'//授業時間数


				'//授業時間数を取得
'				w_bRet = gf_SouJugyo(w_lJikan,w_Rs("T16_KAMOKU_CD"),m_sGakunen,m_sClassNo,w_sKaisibi,w_sSyuryobi,m_iSyoriNen)
'				if w_bRet <> True Then
'					m_AryResult(3,i) = ""					'//授業時間数取得失敗
'				Else
'					m_AryResult(3,i) = w_lJikan				'//授業時間数
'				End If

				m_AryResult(4,i) = w_Rs("KEKA")				'//欠課数
				m_AryResult(5,i) = w_Rs("CHIKAI")			'//遅刻回数

                if m_SchoolNo = "11" then  '福島高専の場合

					Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

					  		Case C_SIKEN_ZEN_TYU    '前期中間試験
				  				  m_AryResult(3,i) = w_Rs("JIKAN")			'//授業時間数
			     				  m_AryResult(4,i) = w_Rs("KEKA")	    	'//欠課数
			     	              m_AryResult(5,i) = w_Rs("CHIKAI") 		'//遅刻回数

							Case C_SIKEN_ZEN_KIM    '前期期末試験

				  				  m_AryResult(3,i) = cint(gf_SetNull2Zero(w_Rs("JIKAN")))	+ cint(gf_SetNull2Zero(w_Rs("JIKAN2")))		'//授業時間数
			     				  m_AryResult(4,i) = cint(gf_SetNull2Zero(w_Rs("KEKA")))	+ cint(gf_SetNull2Zero(w_Rs("KEKA2")))     	'//欠課数
			     	              m_AryResult(5,i) = cint(gf_SetNull2Zero(w_Rs("CHIKAI"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI2")))		'//遅刻回数

							Case C_SIKEN_KOU_TYU    '後期中間試験
				  				  m_AryResult(3,i) = cint(gf_SetNull2Zero(w_Rs("JIKAN")))	+ cint(gf_SetNull2Zero(w_Rs("JIKAN2"))) + cint(gf_SetNull2Zero(w_Rs("JIKAN3")))						'//授業時間数
			     				  m_AryResult(4,i) = cint(gf_SetNull2Zero(w_Rs("KEKA")))	+ cint(gf_SetNull2Zero(w_Rs("KEKA2")))   + cint(gf_SetNull2Zero(w_Rs("KEKA3")))     					'//欠課数
			     	              m_AryResult(5,i) = cint(gf_SetNull2Zero(w_Rs("CHIKAI"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI2"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI3")))						'//遅刻回数

							Case C_SIKEN_KOU_KIM    '後期期末試験
				  				  m_AryResult(3,i) = cint(gf_SetNull2Zero(w_Rs("JIKAN")))	+ cint(gf_SetNull2Zero(w_Rs("JIKAN2"))) + cint(gf_SetNull2Zero(w_Rs("JIKAN3"))) + cint(gf_SetNull2Zero(w_Rs("JIKAN4")))		'//授業時間数
			     				  m_AryResult(4,i) = cint(gf_SetNull2Zero(w_Rs("KEKA")))	+ cint(gf_SetNull2Zero(w_Rs("KEKA2")))   + cint(gf_SetNull2Zero(w_Rs("KEKA3")))   + cint(gf_SetNull2Zero(w_Rs("KEKA4")))   	'//欠課数
			     	              m_AryResult(5,i) = cint(gf_SetNull2Zero(w_Rs("CHIKAI"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI2"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI3"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI4")))	'//遅刻回数
					End Select
                END IF
				i = i + 1
			End If

			w_Rs.MoveNext
		Loop

		m_iCnt = i-1

        '//正常終了
        f_GetResultData = 0


        Exit Do
    Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  科目の開設時期を取得
'*  [引数]  なし
'*  [戻値]  True：開設あり、False：開設なし
'*  [説明]  
'********************************************************************************
Function f_GetKaisetu(p_sKamokuCd)
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetKaisetu = False

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_KAISETU" & m_sGakunen
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "        T15_KAMOKU_CD = '" & p_sKamokuCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    AND T15_NYUNENDO = " & m_iSyoriNen - m_sGakunen + 1
		w_sSQL = w_sSQL & vbCrLf & "    AND ( ("

		Select Case cint(m_sSikenKbn) '選んだ試験によって、取得科目の開設期間を変える
			Case cint(C_SIKEN_ZEN_TYU)
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " "

			Case cint(C_SIKEN_ZEN_KIM)
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " "

			Case cint(C_SIKEN_KOU_TYU)
				'w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " OR "
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_KOUKI & " "

			Case cint(C_SIKEN_KOU_KIM)
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " OR "
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_KOUKI & " "

		End Select

		w_sSQL = w_sSQL & vbCrLf & "    )"
		w_sSQL = w_sSQL & vbCrLf & "    OR ("
		w_sSQL = w_sSQL & vbCrLf & "       T15_KAISETU" & m_sGakunen & "=" & C_KAI_TUNEN & " "
		w_sSQL = w_sSQL & vbCrLf & "    ) ) "

'response.write w_ssql & "<br>"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
		    Call gf_closeObject(rs)
			Exit Function
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetKaisetu = True

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function


'********************************************************************************
'*  [機能]  学科の略称を取得
'*  [引数]  p_sGakkaCd : 学科CD
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetGakkaNm(p_iGakunen,p_iClass)
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKAMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_CD = M05_CLASS.M05_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_NENDO = M05_CLASS.M05_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & p_iGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & p_iClass

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKAMEI")
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetGakkaNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*  [機能]  席次、担任所見を取得
'*  [引数]  p_sGakkaCd : 学科CD
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetGakuseiInfo()
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakuseiInfo = 1
	m_iSikiji = ""
	m_iSyoken = ""
	m_iAverage = 0
	
	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "

		'//試験区分により場合わけ
		Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

			Case C_SIKEN_ZEN_TYU    '前期中間試験
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI_TYUKAN_Z AS SEKIJI"			'//席次
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_NINZU_TYUKAN_Z AS NINZU"			'//クラス人数
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_TYUKAN_Z AS SYOKEN "			'//所見
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_TYUKAN_Z AS HEIKIN "			'//平均点
			Case C_SIKEN_ZEN_KIM    '前期期末試験
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI_KIMATU_Z AS SEKIJI"			'//席次
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_NINZU_KIMATU_Z AS NINZU"			'//クラス人数
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_KIMATU_Z AS SYOKEN "			'//所見
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_KIMATU_Z AS HEIKIN "			'//平均点
			Case C_SIKEN_KOU_TYU    '後期中間試験
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI_TYUKAN_K  AS SEKIJI"			'//席次
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_NINZU_TYUKAN_K  AS NINZU"			'//クラス人数
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_TYUKAN_K AS SYOKEN "			'//所見
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_TYUKAN_K AS HEIKIN "			'//平均点
			Case C_SIKEN_KOU_KIM    '後期期末試験
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI  AS SEKIJI"					'//席次
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASSNINZU  AS NINZU"				'//クラス人数
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_KIMATU_K AS SYOKEN"			'//所見
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_KIMATU_K AS HEIKIN "			'//平均点
			Case Else
				'//システムエラー
	            m_sErrMsg = "試験情報がありません。"
				Exit Do
		End Select

		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T13_GAKU_NEN.T13_GAKUSEI_NO='" & m_sGakusei & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			f_GetGakuseiInfo = 99
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			m_iSikiji = rs("SEKIJI")
			m_iMemberCnt = rs("NINZU")
			m_iSyoken = rs("SYOKEN")
			m_iAverage = rs("HEIKIN")
		End If 

		f_GetGakuseiInfo = 0
		Exit do 
	Loop

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDisp_Data_Siken()
    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_sSikenName

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = ""
	w_sSikenName = ""

    Do

        '試験マスタよりデータを取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_sSikenKBN

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDisp_Data_Siken = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sSikenName = rs("M01_SYOBUNRUIMEI")
        End If

        Exit Do
    Loop

	'//戻り値セット
    f_GetDisp_Data_Siken = w_sSikenName

    Call gf_closeObject(rs)

End Function

'*******************************************************************************
' 機　　能：科目評価取得
' 返　　値：
' 引　　数：p_iTensu - 点数(IN)
'           
'
' 機能詳細：点数から評価NOのp_uDataに評価、評定、欠点科目を設定する
' 備　　考：評価NOがすでに分かっている場合には直接callも可
'           評価NOが分からないときは、gf_GetKamokuTensuHyokaをcall
'           2002.06.19 松尾
'*******************************************************************************
Function f_GetTensuHyoka(p_iTensu)
    Dim w_oRecord
    Dim w_sSql
    
    Const C_HYOKA_FUKA = 1
    
    On Error Resume Next
    
    f_GetTensuHyoka = ""
	
	if gf_SetNull2String(p_iTensu) = "" then exit function
	
    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_RYAKU "
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 	M08_MIN <= " & p_iTensu								'点数
    w_sSql = w_sSql & " AND M08_MAX >= " & p_iTensu
    w_sSql = w_sSql & " AND M08_NENDO = " & m_iSyoriNen							'年度
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_IPPAN		'一般学科
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '科目Mない時エラー
    if w_oRecord.EOF Then Exit Function
    
    'データセット
    if cint(gf_SetNull2Zero(w_oRecord("M08_HYOKA_SYOBUNRUI_RYAKU"))) = C_HYOKA_FUKA then
    	f_GetTensuHyoka = "*"
    end if
    
    Call gf_closeObject(w_oRecord)
    
End Function


Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Dim w_iTaniKei
	Dim w_iHyokaKei
	Dim w_iJikanKei
	Dim w_iKekkaKei
	Dim w_iTikokuKei
	Dim w_iSeiAverage

%>

	<html>

	<head>
	<title>個人別成績一覧</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<SCRIPT language="JavaScript">
	<!--
    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Cansel(){

        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  前へ,次へボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_NextPage(p_FLG){

		if( p_FLG == 1){
			document.frm.txtGakusei.value = document.frm.txtBeforGakuNo.value;
		}else{
        	document.frm.txtGakusei.value = document.frm.txtAfterGakuNo.value;
		}

		document.frm.action="sei0300_main.asp";
		document.frm.target="_self";
		document.frm.submit();
    }

	//-->
	</SCRIPT>


	<link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>
	<body>
	<center>
	<form name="frm" METHOD="post">
	<% call gs_title(" 個人別成績一覧 "," 一　覧 ") %>
	<BR>

	<%Do %>
		<table border="0" width="500" class=hyo align="center">
			<tr>
				<th width="500" class="header" colspan="4"><%=f_GetDisp_Data_Siken()%><br></th>
			</tr>
			<tr>
				<th width="50" class="header2">クラス</th>
	<%If m_sKengen <> C_KENGEN_SEI0300_GAK then%>
				<td width="150" align="center" class="detail"><%=m_sGakunen%>-<%=m_sClassNo%> [<%=f_GetGakkaNm(m_sGakunen,m_sClassNo)%>]</td>
	<%Else%>
				<td width="150" align="center" class="detail"><%=m_sGakunen%>年　<%=gf_GetGakkaNm(m_iSyoriNen,m_sGakkaNo)%></td>
	<%End If%>
				<th width="50" class="header2">氏　名</th>
				<td width="250" align="left" class="detail">　( <%=m_sGakusekiNo%> )<%=m_sName%></td>
			</tr>
		</table>

		<!--ボタン-->
		<BR>
		<table border="0" width="250">
		    <tr>
		        <td valign="top" align="center">
		            <input type="button" value="　前　へ　" class="button" <%If m_sBeforGakuNo = "" Then %> DISABLED <%End If%>  onclick="javascript:f_NextPage(1)">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="キャンセル" class="button" onclick="javascript:f_Cansel()">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="　次　へ　" class="button" <%If m_sAfterGakuNo = "" Then%>  DISABLED <%End If%>  onclick="javascript:f_NextPage(2)">
		        </td>
		    </tr>
		</table>
		<br>

		<!--担任所見-->
		<table class="hyo" border="1" align="center" width="70%">
		<tr>
			<th class="header" colspan="6">担　任　所　見</th>
		</tr>
		<tr>
			<td class="detail" colspan="6" height="35" valign="top"><%=m_iSyoken%><br></td>
		</tr>

		<!--明細部-->
		<tr>
			<th class="header" rowspan="2" width="31%">科　目　名</th>
			<th class="header" rowspan="2" width="13%">単位数</th>
			<th class="header" rowspan="2" width="13%">成績<br>評価</th>
			<th class="header" rowspan="2" width="13%">授業<br>時間数</th>
			<th class="header" colspan="2" width="20%">出　席　状　況</th>
		</tr>
		<tr>
			<th class="header">欠課<br>時間</th>
			<th class="header">遅刻<br>回数</th>
		</tr>

		<% 
		'//合計初期化
		w_iTaniKei   = 0
		w_iHyokaKei  = 0
		w_iJikanKei  = 0
		w_iKekkaKei  = 0
		w_iTikokuKei = 0

		For i = 0 To m_iCnt

			call gs_cellPtn(w_cell) %>
			<tr>
				<td class="<%=w_cell%>" align="left" ><%=m_AryResult(0,i)%><br></td>
				<td class="<%=w_cell%>" align="right"><%=FormatNumber(cint(gf_SetNull2Zero(m_AryResult(1,i))),1)%><br></td>

				<% If Cint(m_AryResult(6,i)) = 0 Then %>

					<td class="<%=w_cell%>" align="right"><%=f_GetTensuHyoka(m_AryResult(2,i))%>　<%=m_AryResult(2,i)%></td>

				<% ELSE %>

					<td class="<%=w_cell%>" align="right">　<%=m_AryResult(2,i)%></td>

				<% END IF%>
				<td class="<%=w_cell%>" align="right"><%=gf_SetNull2Zero(m_AryResult(3,i))%><br></td>
				<td class="<%=w_cell%>" align="right"><%=gf_SetNull2Zero(m_AryResult(4,i))%><br></td>
				<td class="<%=w_cell%>" align="right"><%=gf_SetNull2Zero(m_AryResult(5,i))%><br></td>
			</tr>

			<%
			'//単位数合計
			w_iTaniKei = w_iTaniKei + cint(gf_SetNull2Zero(m_AryResult(1,i)))

			'//成績評価合計
			If Cint(m_AryResult(6,i)) = 0 Then
				w_iHyokaKei = w_iHyokaKei + cint(gf_SetNull2Zero(m_AryResult(2,i)))
			End if

			'//授業時間数合計
			w_iJikanKei = w_iJikanKei + cint(gf_SetNull2Zero(m_AryResult(3,i)))

			'//欠課数合計
			w_iKekkaKei = w_iKekkaKei + cint(gf_SetNull2Zero(m_AryResult(4,i)))

			'//遅刻回数合計
			w_iTikokuKei = w_iTikokuKei + cint(gf_SetNull2Zero(m_AryResult(5,i)))
			%>

		<% next %>

		<!--合計-->
		<tr>
			<td class="NOCHANGE">合計</td>
			<td class="NOCHANGE" align="right"><%=FormatNumber(w_iTaniKei,1)%></td>
			<td class="NOCHANGE" align="right"><%=w_iHyokaKei%></td>
			<td class="NOCHANGE" align="right"><%=w_iJikanKei%></td>
			<td class="NOCHANGE" align="right"><%=w_iKekkaKei%></td>
			<td class="NOCHANGE" align="right"><%=w_iTikokuKei%></td>
		</tr>
		
		<!--平均-->
		<tr>
			<td class="CELL2">平均</td>
			<td class="CELL2" align="right">―</td>
			<td class="CELL2" align="right"><%=m_iAverage%></td>
			<td class="CELL2" colspan="3" align="right">席次　<%=m_iSikiji%>位／<%=m_iMemberCnt%>人中</td>
		</tr>
		</table>

		<!--ボタン-->
		<BR>
		<table border="0" width="250">
		    <tr>
		        <td valign="top" align="center">
		            <input type="button" value="　前　へ　" class="button" <%If m_sBeforGakuNo = "" Then %> DISABLED <%End If%>  onclick="javascript:f_NextPage(1)">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="キャンセル" class="button" onclick="javascript:f_Cansel()">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="　次　へ　" class="button" <%If m_sAfterGakuNo = "" Then%>  DISABLED <%End If%>  onclick="javascript:f_NextPage(2)">
		        </td>
		    </tr>
		</table>

		<%Exit Do%>
	<%Loop%>

	<input type="hidden" name="txtSikenKBN" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"   value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
	<input type="hidden" name="txtGakkaNo"  value="<%=m_sGakkaNo%>">
	<input type="hidden" name="txtGakusei"  value="<%=m_sGakusei%>">

	<input type="hidden" name="txtBeforGakuNo" value="<%=m_sBeforGakuNo%>">
	<input type="hidden" name="txtAfterGakuNo" value="<%=m_sAfterGakuNo%>">

	</form>
	</center>
	</body>
	</html>
<%
    '---------- HTML END   ----------
End Sub
%>
